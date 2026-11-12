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

function createWorksheetRecord(e) {
  return !e.trim() || e.length > 31 || /[\\/\[\]*?:]/.test(e)
    ? throwWithMessage(
        "Sheet names must contain 1–31 characters and cannot contain \\ / [ ] * ? : .",
      )
    : {
        id: crypto.randomUUID(),
        name: e.trim(),
        cells: {},
        rowCount: workbookLimits.maxRows,
        colCount: workbookLimits.maxCols,
        frozenRows: 0,
      };
}
function writeAnalysisOutputBlock(e, r, t, destination) {
  destination?.sheetId || assertWorkbookStructureEditable(e);
  const o = t[0]?.length ?? 0;
  ((!o || t.length * o > workbookLimits.maxCells) &&
    throwWorkbookError("The analysis output exceeds 100,000 cells."),
    t.some((g) => g.some((b) => typeof b == "string" && b.length > 32767)) &&
      throwWorkbookError(
        "An analysis label exceeds the 32,767-character cell limit. Shorten source labels or value-field names.",
      ));
  const workbook = structuredClone(e),
    sheet = destination?.sheetId
      ? findSheetOrThrow(workbook, destination.sheetId)
      : createWorksheetRecord(
          buildUniqueSheetName(workbook, destination?.sheetName ?? r),
        );
  (destination?.sheetId || workbook.sheets.push(sheet),
    assertSheetUnprotected(sheet));
  const d = _t(destination?.address ?? "A1"),
    lastRow = d.row + t.length - 1,
    lastCol = d.col + o - 1;
  (lastRow >= sheet.rowCount || lastCol >= sheet.colCount) &&
    throwWorkbookError("The output does not fit the destination worksheet.");
  const c = {
    firstRow: d.row,
    lastRow: lastRow,
    firstCol: d.col,
    lastCol: lastCol,
  };
  for (const cell of Object.keys(sheet.cells)) {
    const b = _t(cell);
    Dp(c, b.row, b.col) &&
      throwWorkbookError(
        "The output destination contains existing cells. Choose an empty destination or a new sheet.",
      );
  }
  return (
    sheet.tables?.some((table) =>
      rangesIntersect(c, parseRangeAddress(table.range)),
    ) && throwWorkbookError("The output destination overlaps a table."),
    (sheet.merges?.some((merge) =>
      rangesIntersect(c, parseRangeAddress(merge)),
    ) ||
      sheet.pivots?.some((pivot) =>
        rangesIntersect(c, parseRangeAddress(pivot.outputRange)),
      ) ||
      sheet.queries?.some((query) =>
        rangesIntersect(c, parseRangeAddress(query.outputRange)),
      )) &&
      throwWorkbookError(
        "The output destination overlaps an existing merged area or analysis result. Choose another empty destination.",
      ),
    t.forEach((g, b) =>
      g.forEach((raw, w) => {
        sheet.cells[Cn(d.row + b, d.col + w)] = {
          raw: raw,
          ...(typeof raw == "string" && raw.startsWith("=")
            ? { literal: true }
            : {}),
          ...(b === 0
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
      range: `${Cn(d.row, d.col)}:${Cn(lastRow, lastCol)}`,
    }
  );
}
function buildPivotTableOutput(workbook, r) {
  const pivotDefinition = normalizePivotDefinition(r),
    pivotSource = buildPivotSource(workbook, pivotDefinition.source),
    rows = filterPivotSourceRows(
      pivotDefinition,
      pivotSource.fields,
      filterSourceRows(pivotSource, pivotDefinition.filters ?? []),
    ),
    pivotTable = computePivotTable(pivotDefinition, pivotSource.fields, rows),
    s = writeAnalysisOutputBlock(
      workbook,
      pivotDefinition.name,
      pivotTable.matrix,
      pivotDefinition.destination,
    ),
    { destination, ...u } = pivotDefinition,
    l = _t(pivotDefinition.destination?.address ?? "A1");
  for (let f = 1; f < pivotTable.matrix.length; f++)
    for (
      let p = Math.max(1, pivotDefinition.rowFields.length);
      p < pivotTable.matrix[0].length;
      p++
    ) {
      const m =
        pivotDefinition.values[
          (p - Math.max(1, pivotDefinition.rowFields.length)) %
            pivotDefinition.values.length
        ];
      m.showAs &&
        m.showAs !== "value" &&
        (s.sheet.cells[Cn(l.row + f, l.col + p)].format = "percent");
    }
  if (!pivotDefinition.destination?.sheetId) {
    let f = 26;
    (pivotTable.matrix[0].forEach((p, m) => {
      const v = String(p ?? ""),
        g = Math.min(320, Math.max(90, v.length * 7 + 24));
      s.sheet.columnWidths = { ...s.sheet.columnWidths, [l.col + m]: g };
      const b = Cn(l.row, l.col + m),
        y = s.sheet.cells[b];
      ((y.style = { ...y.style, wrap: true }),
        (f = Math.max(
          f,
          Math.min(120, Math.ceil((v.length * 7) / (g - 16)) * 18 + 8),
        )));
    }),
      (s.sheet.rowHeights = { ...s.sheet.rowHeights, [l.row]: f }));
  }
  for (const subtotalRow of pivotTable.subtotalRows)
    for (let p = 0; p < pivotTable.matrix[0].length; p++) {
      const m = Cn(l.row + subtotalRow, l.col + p),
        v = s.sheet.cells[m];
      v.style = { ...v.style, bold: true, fill: "#e4edf7" };
    }
  return (
    (s.sheet.pivots = [
      ...(s.sheet.pivots ?? []),
      {
        ...u,
        source: pivotSource.source,
        id: createUuid(),
        outputRange: s.range,
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
          outputFingerprint: collectRangeValues(s.sheet, s.range),
          sourceRows: pivotSource.rows.length,
          includedRows: rows.length,
          excludedRows: pivotSource.rows.length - rows.length,
          measures: pivotTable.measures,
        },
      },
    ]),
    s.workbook
  );
}
function refreshPivotTable(workbook, pivotId, destination) {
  const n =
    workbook.sheets
      .flatMap((sheet) => sheet.pivots ?? [])
      .find((o) => o.id === pivotId) ??
    throwWorkbookError("Pivot definition not found.");
  return buildPivotTableOutput(workbook, {
    ...n,
    destination: destination,
    name: `${n.name.slice(0, 90)} refreshed`,
  });
}
function applyPivotViewChange(workbook, pivotId, t, n) {
  (n.readOnly && throwWorkbookError("This workbook is read-only."),
    n.snapshot !== JSON.stringify(workbook) &&
      throwWorkbookError("The workbook changed. Preview pivot filters again."),
    (!isPlainDataObject(t) ||
      Object.keys(t).some(
        (s) =>
          !["slicers", "timelines", "sorts", "values", "chart"].includes(s),
      )) &&
      throwWorkbookError(
        "Pivot view changes can only set filters, sorting, value display and chart settings.",
      ));
  const { owner, definition } = findPivotDefinition(workbook, pivotId);
  return (
    assertSheetUnprotected(owner),
    buildPivotTableOutput(workbook, {
      ...definition,
      ...t,
      name: `${definition.name.slice(0, 88)} filtered`,
      destination: undefined,
    })
  );
}
function extractPivotDrillThrough(workbook, pivotId, t, n) {
  (n.readOnly && throwWorkbookError("This workbook is read-only."),
    n.snapshot !== JSON.stringify(workbook) &&
      throwWorkbookError("The workbook changed. Preview drill-through again."));
  const pivotDrillDown = buildPivotDrillDown(workbook, pivotId, t),
    { owner } = findPivotDefinition(workbook, pivotId);
  assertSheetUnprotected(owner);
  const s = JSON.stringify({ ...pivotDrillDown, records: undefined });
  s.length > 32700 &&
    throwWorkbookError(
      "Drill context exceeds the cell-note budget. Reduce filter selections before extracting.",
    );
  const d = [
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
    u = writeAnalysisOutputBlock(
      workbook,
      `${pivotDrillDown.context.name.slice(0, 18)} details`,
      d,
    );
  return (
    (u.sheet.cells.A1.note = `Pivot drill-through context (snapshot, not signed): ${s}`),
    (u.sheet.columnWidths = { 0: 240, 1: 110, 2: 160, 3: 130, 4: 110 }),
    (u.sheet.rowOriginIds = {}),
    (u.workbook.rowOrigins = [...(workbook.rowOrigins ?? [])]),
    pivotDrillDown.records.forEach((record, c) => {
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
        u.workbook.rowOrigins.length >= workbookLimits.maxCells &&
          throwWorkbookError(
            "Row provenance exceeds its 100,000 identity budget.",
          ),
        u.workbook.rowOrigins.push(original)),
        (u.sheet.rowOriginIds[c + 1] = original.id),
        record.cells.forEach(({ cell }, p) => {
          const m = [
            cell.note,
            typeof cell.raw == "string" &&
            cell.raw.startsWith("=") &&
            !cell.literal
              ? `Source formula at extraction: ${cell.raw}`
              : undefined,
          ].filter(Boolean).join(`
`);
          (m.length > 32767 &&
            throwWorkbookError(
              "A source note/formula exceeds the drill-output note budget.",
            ),
            m &&
              (u.sheet.cells[`${columnIndexToLetters(p + 5)}${c + 2}`].note =
                m));
        }));
    }),
    u.workbook.sheets.reduce(
      (l, c) => l + Object.keys(c.rowOriginIds ?? {}).length,
      0,
    ) > workbookLimits.maxCells &&
      throwWorkbookError(
        "Drill row associations exceed the 100,000 source-row budget.",
      ),
    u.workbook
  );
}
function runQueryDefinition(workbook, r) {
  const queryDefinition = normalizeQueryDefinition(r),
    pivotSource = buildPivotSource(workbook, queryDefinition.source),
    o = filterSourceRows(pivotSource, queryDefinition.filters ?? []),
    a = queryDefinition.fields.map((field) =>
      findSourceFieldIndex(pivotSource.fields, field),
    ),
    s = (queryDefinition.derived ?? []).map((f) => ({
      ...f,
      index: findSourceFieldIndex(pivotSource.fields, f.field),
      other: f.otherField
        ? findSourceFieldIndex(pivotSource.fields, f.otherField)
        : undefined,
    }));
  if (queryDefinition.sort) {
    const sourceFieldIndex = findSourceFieldIndex(
        pivotSource.fields,
        queryDefinition.sort.field,
      ),
      p = queryDefinition.sort.direction === "ascending" ? 1 : -1;
    o.sort((m, v) => {
      const g = m.values[sourceFieldIndex],
        b = v.values[sourceFieldIndex];
      return (
        (isFiniteNumber(g) && isFiniteNumber(b)
          ? g - b
          : String(g ?? "").localeCompare(String(b ?? ""))) * p ||
        m.index - v.index
      );
    });
  }
  const d = [
      ...queryDefinition.fields.map((field) =>
        Object.hasOwn(queryDefinition.rename ?? {}, field)
          ? queryDefinition.rename[field]
          : field,
      ),
      ...s.map((f) => f.name),
    ],
    u = [d];
  (o.length + 1) * d.length > workbookLimits.maxCells &&
    throwWorkbookError("The query output exceeds 100,000 cells.");
  for (const f of o) {
    const p = a.map((m) => f.values[m]);
    for (const m of s) {
      const v = f.values[m.index],
        g = m.other !== undefined ? f.values[m.other] : (m.constant ?? null);
      if (m.operation === "upper") p.push(String(v ?? "").toUpperCase());
      else if (m.operation === "lower") p.push(String(v ?? "").toLowerCase());
      else if (m.operation === "trim") p.push(String(v ?? "").trim());
      else if (m.operation === "concat")
        p.push(String(v ?? "") + String(g ?? ""));
      else {
        ((!isFiniteNumber(v) || !isFiniteNumber(g)) &&
          throwWorkbookError(
            `Derived field “${m.name}” requires numbers on source row ${f.index + 1}.`,
          ),
          m.operation === "divide" &&
            g === 0 &&
            throwWorkbookError(
              `Derived field “${m.name}” divides by zero on source row ${f.index + 1}.`,
            ));
        const b =
          m.operation === "add"
            ? v + g
            : m.operation === "subtract"
              ? v - g
              : m.operation === "multiply"
                ? v * g
                : v / g;
        (Number.isFinite(b) ||
          throwWorkbookError(
            "Derived calculation exceeds finite numeric bounds.",
          ),
          p.push(b));
      }
    }
    (p.some((m) => typeof m == "string" && m.length > 32767) &&
      throwWorkbookError("Derived text exceeds the cell text limit."),
      u.push(p));
  }
  const l = writeAnalysisOutputBlock(
      workbook,
      queryDefinition.name,
      u,
      queryDefinition.destination,
    ),
    { destination, ...h } = queryDefinition;
  return (
    (l.sheet.queries = [
      ...(l.sheet.queries ?? []),
      {
        ...h,
        source: pivotSource.source,
        id: createUuid(),
        sourceRevision: workbook.revision,
        outputRange: l.range,
      },
    ]),
    l.workbook
  );
}
function refreshQueryDefinition(workbook, pivotId, destination) {
  const n =
    workbook.sheets
      .flatMap((sheet) => sheet.queries ?? [])
      .find((o) => o.id === pivotId) ??
    throwWorkbookError("Query definition not found.");
  return runQueryDefinition(workbook, {
    ...n,
    destination: destination,
    name: `${n.name.slice(0, 90)} refreshed`,
  });
}
function normalizeCellRecord(cell, r) {
  if (
    r &&
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
  const n = { raw: raw };
  if (cell.literal !== undefined) {
    if (typeof cell.literal != "boolean")
      return throwWithMessage("Cell literal flag must be boolean.");
    n.literal = cell.literal;
  }
  if (
    (cell.numberFormat !== undefined &&
      (n.numberFormat = validateNumberFormatSettings(cell.numberFormat)),
    cell.numFmt !== undefined &&
      (n.numFmt = requireBoundedText(cell.numFmt, "Excel number format", 200)),
    cell.format !== undefined)
  ) {
    if (
      typeof cell.format != "string" ||
      !numberFormatCategories.has(cell.format)
    )
      return throwWithMessage("Unsupported cell format.");
    n.format = cell.format;
  }
  return (
    cell.hyperlink !== undefined &&
      (n.hyperlink = validateHyperlinkRecord(cell.hyperlink)),
    cell.note !== undefined &&
      (n.note = requireBoundedText(cell.note, "Cell note", 1e4)),
    cell.tickmark !== undefined &&
      (n.tickmark = requireBoundedText(cell.tickmark, "Tickmark", 40)),
    cell.style !== undefined && (n.style = validateCellStyle(cell.style)),
    n
  );
}
function normalizeCellRecordInput(cell) {
  return normalizeCellRecord(cell, false);
}
function paintCellFormatting(workbook, sheetId, t, n) {
  const sheet = workbook.sheets.find((sheet) => sheet.id === sheetId);
  assertSheetUnprotected(sheet);
  const a = assertRangeCellBudget(n);
  a.length > 500 && failWithMessage("Paint at most 500 cells.");
  const s = sheet.cells[t] ?? {},
    cells = { ...sheet.cells };
  for (const l of a)
    cells[l] = normalizeCellRecordInput({
      ...cells[l],
      raw: cells[l]?.raw ?? null,
      style: structuredClone(s.style ?? {}),
      format: s.format ?? "general",
      numFmt: s.numFmt,
      numberFormat: s.numberFormat
        ? structuredClone(s.numberFormat)
        : undefined,
    });
  return {
    ...workbook,
    sheets: workbook.sheets.map((sheet) =>
      sheet.id === sheetId ? { ...sheet, cells: cells } : sheet,
    ),
  };
}
function validateTableSourceEvidence(tableEvidence, r) {
  if (tableEvidence === undefined) return;
  const t = () => {
      throw Error(
        "Invalid or over-capacity table source evidence. Keep the native backup.",
      );
    },
    n = (s) => !!s && typeof s == "object" && !Array.isArray(s);
  (!Array.isArray(tableEvidence) ||
    tableEvidence.length > 100 ||
    JSON.stringify(tableEvidence).length > 5e6) &&
    t();
  const o = new Set();
  let a = 0;
  r.tableEvidence = tableEvidence.map((s) => {
    (!n(s) ||
      Object.keys(s).some(
        (c) =>
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
          ].includes(c),
      )) &&
      t();
    const d = s;
    for (const c of ["id", "sheetId", "sheetName", "tableId", "tableName"])
      (typeof d[c] != "string" || !d[c] || d[c].length > 512) && t();
    ((o.has(d.id) ||
      !Number.isSafeInteger(d.sourceRevision) ||
      Number(d.sourceRevision) < 0 ||
      !["headers", "columnType"].includes(String(d.kind)) ||
      !Array.isArray(d.changes) ||
      d.changes.length > 500 ||
      (a += d.changes.length) > 1e4) &&
      t(),
      d.kind === "columnType"
        ? ((typeof d.columnId != "string" ||
            !d.columnId ||
            d.columnId.length > 100) &&
            t(),
          validateTableColumnType(d.dataType))
        : (d.columnId !== undefined || d.dataType !== undefined) && t(),
      o.add(d.id));
    const u = new Set(),
      changes = d.changes.map((change) => {
        (!n(change) ||
          Object.keys(change).some(
            (m) => !["address", "before", "after"].includes(m),
          ) ||
          typeof change.address != "string") &&
          t();
        const h = change,
          address = h.address,
          p = _t(address);
        return (
          (`${columnIndexToLetters(p.col)}${p.row + 1}` !== address ||
            u.has(address)) &&
            t(),
          u.add(address),
          {
            address: address,
            ...(h.before !== undefined
              ? { before: normalizeCellRecordInput(h.before) }
              : {}),
            after: normalizeCellRecordInput(h.after),
          }
        );
      });
    return { ...d, changes: changes };
  });
}
function previewColumnTypeChange(workbook, sheetId, tableId, request, o = {}) {
  if (o.readOnly) throw Error("This workbook is read-only.");
  if (
    !request ||
    typeof request != "object" ||
    Object.keys(request).some(
      (E) =>
        ![
          "columnId",
          "dataType",
          "mode",
          "invalid",
          "decimal",
          "group",
        ].includes(E),
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
  const a = structuredClone(request),
    sheetOrThrow = findSheetOrThrow(workbook, sheetId),
    d = sheetOrThrow.tables?.find((table) => table.id === tableId);
  if ((assertSheetUnprotected(sheetOrThrow), !d))
    throw Error("Table not found.");
  const u = d.columns.findIndex((column) => column.id === a.columnId);
  if (u < 0) throw Error("Select an existing table column.");
  const rangeAddress = parseRangeAddress(d.range),
    c = rangeAddress.lastRow - rangeAddress.firstRow - Number(d.showTotals);
  if (c > 500)
    throw Error(
      "A reviewed column type operation supports at most 500 data rows. Preserve the native source and use bounded extracts for larger conversions.",
    );
  const next = structuredClone(workbook),
    f = findSheetOrThrow(next, sheetId),
    p = f.tables.find((table) => table.id === tableId).columns[u];
  a.dataType === "any" ? delete p.dataType : (p.dataType = a.dataType);
  const changes = [],
    exceptions = [],
    workbookWithSettings = calculateWorkbookWithSettings(workbook);
  for (let E = 1; E <= c; E++) {
    const address = `${columnIndexToLetters(rangeAddress.firstCol + u)}${rangeAddress.firstRow + E + 1}`,
      cell = sheetOrThrow.cells[address],
      F = cell?.raw ?? null,
      I = typeof F == "string" && F.startsWith("=") && !cell?.literal;
    try {
      if (I || a.mode === "tag") {
        if (
          !coerceImportedValue(
            workbookWithSettings(sheetId, address),
            a.dataType,
          )
        )
          throw Error(
            I
              ? "Formula result does not match the selected type; the formula is preserved."
              : "Value does not match the selected type; source values are unchanged.",
          );
        continue;
      }
      const O = workbookWithSettings(sheetId, address);
      if (isNullOrEmptyValue(F) && !isNullOrEmptyValue(O)) {
        if (!coerceImportedValue(O, a.dataType))
          throw Error(
            "Spill result does not match the selected type; the spill is preserved.",
          );
        continue;
      }
      const raw = convertTextForDataType(F, a);
      if (raw !== F) {
        const cellRecordInput = normalizeCellRecordInput({
          ...cell,
          raw: raw,
          ...(typeof raw == "string" && raw.startsWith("=")
            ? { literal: true }
            : { literal: undefined }),
        });
        ((f.cells[address] = cellRecordInput),
          changes.push({
            address: address,
            ...(cell ? { before: structuredClone(cell) } : {}),
            after: structuredClone(cellRecordInput),
          }));
      }
    } catch (error) {
      exceptions.push({ address: address, value: F, reason: error.message });
    }
  }
  const beforeType = d.columns[u].dataType ?? "any";
  return (
    (changes.length || beforeType !== a.dataType) &&
      ((next.tableEvidence = [
        ...(next.tableEvidence ?? []),
        {
          id: `${tableId}:${p.id}:${workbook.revision}:${workbook.tableEvidence?.length ?? 0}`,
          kind: "columnType",
          sourceRevision: workbook.revision,
          sheetId: sheetId,
          sheetName: sheetOrThrow.name,
          tableId: tableId,
          tableName: d.name,
          columnId: p.id,
          dataType: a.dataType,
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
      request: a,
      beforeType: beforeType,
      canApply: !exceptions.length || a.invalid === "keep",
      changes: changes,
      exceptions: exceptions,
      next: next,
    }
  );
}
function applyColumnTypePreview(workbook, plan, t = {}) {
  if (JSON.stringify(workbook) !== plan.snapshot)
    throw Error("The workbook changed. Preview column types again.");
  const n = previewColumnTypeChange(
    workbook,
    plan.sheetId,
    plan.tableId,
    plan.request,
    t,
  );
  if (JSON.stringify(n) !== JSON.stringify(plan))
    throw Error("The column type preview changed. Preview again.");
  if (!n.canApply)
    throw Error(
      "Resolve the listed exceptions or explicitly choose to retain them unchanged.",
    );
  return JSON.stringify(workbook) === JSON.stringify(n.next)
    ? workbook
    : n.next;
}
function createWorksheetTable(workbook, sheetId, t) {
  const sheetOrThrow = findSheetOrThrow(workbook, sheetId),
    tableName = validateTableName(t.name),
    sourceRange = resolveSourceRange(t.range, sheetOrThrow);
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
    t.showTotals)
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
  const s = structuredClone(workbook),
    d = findSheetOrThrow(s, sheetId),
    u = {
      id: createUuid(),
      name: tableName,
      range: `${Cn(sourceRange.firstRow, sourceRange.firstCol)}:${Cn(sourceRange.lastRow, sourceRange.lastCol)}`,
      columns: validateReviewedTableHeaders(
        workbook,
        sheetId,
        t.range,
        t.headers,
      ).map((f) => ({ id: createUuid(), name: f })),
      style: t.style ?? "blue",
      showTotals: !!t.showTotals,
      totals: t.totals ?? {},
    };
  ((d.tables = [...(d.tables ?? []), u]),
    normalizeTableDefinitions(d, d),
    applyTableStyleToCells(d, d.tables[d.tables.length - 1]));
  const changes = u.columns.flatMap((column, p) => {
    const address = Cn(sourceRange.firstRow, sourceRange.firstCol + p),
      cell = sheetOrThrow.cells[address],
      g = d.cells[address];
    return cell?.raw === column.name &&
      !(column.name.startsWith("=") && !cell?.literal)
      ? []
      : [
          {
            address: address,
            ...(cell ? { before: structuredClone(cell) } : {}),
            after: structuredClone(g),
          },
        ];
  });
  return (
    changes.length &&
      ((s.tableEvidence = [
        ...(s.tableEvidence ?? []),
        {
          id: `${u.id}:headers`,
          kind: "headers",
          sourceRevision: workbook.revision,
          sheetId: sheetId,
          sheetName: sheetOrThrow.name,
          tableId: u.id,
          tableName: tableName,
          changes: changes,
        },
      ]),
      validateTableSourceEvidence(s.tableEvidence, s)),
    assertWithinCellBudget(s)
  );
}
function fillCalculatedTableColumns(e, r, firstRow, lastRow) {
  const o = tableDataBodyRange(r);
  for (const [a, s] of r.columns.entries())
    if (s.calculatedFormula)
      for (let d = firstRow; d <= lastRow; d++) {
        const u = Cn(d, o.firstCol + a),
          raw = shiftFormulaReferences(s.calculatedFormula, d - o.firstRow, 0);
        validateCalculatedColumnFormula(raw);
        const cell = { ...(e.cells[u] ?? { raw: null }), raw: raw };
        (delete cell.literal, (e.cells[u] = normalizeCellRecordInput(cell)));
      }
}
function resizeTable(workbook, sheetId, tableId, n) {
  const sheetOrThrow = findSheetOrThrow(workbook, sheetId),
    a =
      sheetOrThrow.tables?.find((table) => table.id === tableId) ??
      throwWorkbookError("Table not found.");
  assertSheetUnprotected(sheetOrThrow);
  const tableName = validateTableName(n.name),
    sourceRange = resolveSourceRange(n.range, sheetOrThrow),
    rangeAddress = parseRangeAddress(a.range);
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
  const showTotals = n.showTotals ?? a.showTotals;
  if (
    (showTotals !== a.showTotals &&
      toA1RangeAddress(n.range) === a.range &&
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
      (!a.showTotals || sourceRange.lastRow !== rangeAddress.lastRow))
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
  const c = a.columns.some((column) => column.calculatedFormula);
  (c &&
    sourceRange.lastCol !== rangeAddress.lastCol &&
    throwWorkbookError(
      "Clear calculated column rules before changing table columns.",
    ),
    c &&
      (sourceRange.lastRow !== rangeAddress.lastRow ||
        showTotals !== a.showTotals) &&
      resizeTableRange(workbook, sheetId, a, sourceRange, showTotals));
  const h = structuredClone(workbook),
    f = findSheetOrThrow(h, sheetId);
  if (
    a.showTotals &&
    (!showTotals || sourceRange.lastRow !== rangeAddress.lastRow)
  )
    for (
      let firstCol = rangeAddress.firstCol;
      firstCol <= rangeAddress.lastCol;
      firstCol++
    ) {
      const k = Cn(rangeAddress.lastRow, firstCol),
        cell = f.cells[k],
        tableTotalCell = buildTableTotalCell(
          a,
          firstCol - rangeAddress.firstCol,
        );
      (cell &&
        cell.raw !== tableTotalCell.raw &&
        throwWorkbookError(
          "The old totals row was edited. Preserve or move those values before resizing it.",
        ),
        delete f.cells[k]);
    }
  const p = validateReviewedTableHeaders(h, sheetId, n.range, n.headers),
    columns = p.map((A, k) => ({
      ...a.columns[k],
      id: a.columns[k]?.id ?? `${a.id}:column:${k}:${workbook.revision}`,
      name: A,
    })),
    totals =
      n.totals ??
      Object.fromEntries(
        Object.entries(a.totals).filter(([A]) =>
          columns.some((column) => column.name === A),
        ),
      ),
    g = {
      ...a,
      name: tableName,
      range: `${Cn(sourceRange.firstRow, sourceRange.firstCol)}:${Cn(sourceRange.lastRow, sourceRange.lastCol)}`,
      columns: columns,
      style: n.style ?? a.style,
      showTotals: showTotals,
      totals: totals,
    };
  ((f.tables = f.tables.map((table) => (table.id === tableId ? g : table))),
    normalizeTableDefinitions(f, f),
    applyTableStyleToCells(f, g));
  const changes = columns.flatMap((column, k) => {
    const address = Cn(sourceRange.firstRow, sourceRange.firstCol + k),
      cell = sheetOrThrow.cells[address],
      O = f.cells[address];
    return cell?.raw === O.raw && !!cell?.literal == !!O.literal
      ? []
      : [
          {
            address: address,
            ...(cell ? { before: structuredClone(cell) } : {}),
            after: structuredClone(O),
          },
        ];
  });
  if (
    (changes.length &&
      ((h.tableEvidence = [
        ...(h.tableEvidence ?? []),
        {
          id: `${tableId}:headers:${workbook.revision}:${h.tableEvidence?.length ?? 0}`,
          kind: "headers",
          sourceRevision: workbook.revision,
          sheetId: sheetId,
          sheetName: sheetOrThrow.name,
          tableId: tableId,
          tableName: tableName,
          changes: changes,
        },
      ]),
      validateTableSourceEvidence(h.tableEvidence, h)),
    c)
  ) {
    const A = tableDataBodyRange(a),
      k = tableDataBodyRange(g);
    for (
      let firstRow = rangeAddress.firstRow;
      sourceRange.lastRow !== rangeAddress.lastRow &&
      firstRow <= Math.min(A.lastRow, k.lastRow);
      firstRow++
    )
      for (let firstCol = A.firstCol; firstCol <= A.lastCol; firstCol++) {
        const O = Cn(firstRow, firstCol);
        sheetOrThrow.cells[O] &&
          (f.cells[O] = structuredClone(sheetOrThrow.cells[O]));
      }
    (a.showTotals &&
      showTotals &&
      sourceRange.lastRow !== rangeAddress.lastRow &&
      moveTotalsRowCells(workbook, h, sheetId, a, g),
      fillCalculatedTableColumns(
        f,
        g,
        Math.max(A.lastRow + 1, k.firstRow),
        k.lastRow,
      ));
  }
  return assertWithinCellBudget(renameTableHeaderField(workbook, h));
}
function applyTableOperation(workbook, sheetId, tableId, request, o = {}) {
  o.readOnly && throwWorkbookError("This workbook is read-only.");
  const sheetOrThrow = findSheetOrThrow(workbook, sheetId),
    s =
      sheetOrThrow.tables?.find((table) => table.id === tableId) ??
      throwWorkbookError("Select an existing table.");
  (assertSheetUnprotected(sheetOrThrow),
    isPlainDataObject(request) ||
      throwWorkbookError("Choose a table operation."));
  const d = {
    setColumnFormula: ["kind", "columnId", "formula"],
    clearColumnFormula: ["kind", "columnId"],
    appendRows: ["kind", "count", "rows"],
    resize: ["kind", "range"],
    updateTable: ["kind", "options"],
  };
  (!d[request.kind] ||
    Object.keys(request).some((v) => !d[request.kind].includes(v))) &&
    throwWorkbookError("Unknown calculated-table operation field.");
  let next = structuredClone(workbook),
    l = findSheetOrThrow(next, sheetId),
    table = l.tables.find((table) => table.id === tableId);
  if (
    request.kind === "setColumnFormula" ||
    request.kind === "clearColumnFormula"
  ) {
    const v =
      table.columns.find((column) => column.id === request.columnId) ??
      throwWorkbookError("Select an existing table column.");
    if (request.kind === "setColumnFormula") {
      v.calculatedFormula = validateCalculatedColumnFormula(request.formula);
      const g = tableDataBodyRange(table);
      (g.lastRow - g.firstRow + 1 > tableMutationLimits.changedCells &&
        throwWorkbookError(
          "A calculated column fill supports at most 500 data-body rows.",
        ),
        fillCalculatedTableColumns(
          l,
          {
            ...table,
            columns: table.columns.map((column) =>
              column.id === v.id
                ? column
                : { ...column, calculatedFormula: undefined },
            ),
          },
          g.firstRow,
          g.lastRow,
        ));
    } else delete v.calculatedFormula;
  } else if (
    request.kind === "appendRows" ||
    request.kind === "resize" ||
    request.kind === "updateTable"
  ) {
    const rangeAddress = parseRangeAddress(s.range),
      g = tableDataBodyRange(s);
    let b,
      y = 0,
      w;
    request.kind === "appendRows"
      ? (request.rows !== undefined
          ? (request.count !== undefined &&
              throwWorkbookError("Specify raw rows or a row count, not both."),
            (!Array.isArray(request.rows) ||
              request.rows.length > tableMutationLimits.appendedRows ||
              request.rows.some(
                (row) => !Array.isArray(row) || row.length !== s.columns.length,
              )) &&
              throwWorkbookError(
                "Each appended row must contain one raw value for every table column.",
              ),
            (b = request.rows),
            (y = b.length))
          : (y = request.count ?? 0),
        (!Number.isInteger(y) ||
          y < 1 ||
          y > tableMutationLimits.appendedRows) &&
          throwWorkbookError("Append 1–100 rows at a time."),
        rangeAddress.lastRow + y >= sheetOrThrow.rowCount &&
          throwWorkbookError(
            "The appended table exceeds the worksheet boundary.",
          ),
        (w = {
          ...s,
          range: `${Cn(rangeAddress.firstRow, rangeAddress.firstCol)}:${Cn(rangeAddress.lastRow + y, rangeAddress.lastCol)}`,
        }))
      : (w =
          request.kind === "resize"
            ? { ...s, range: request.range }
            : request.options);
    const sourceRange = resolveSourceRange(w.range, sheetOrThrow),
      showTotals = w.showTotals ?? s.showTotals;
    if (
      ((request.kind !== "updateTable" ||
        s.columns.some((column) => column.calculatedFormula)) &&
        resizeTableRange(workbook, sheetId, s, sourceRange, showTotals),
      (next = resizeTable(workbook, sheetId, tableId, w)),
      (l = findSheetOrThrow(next, sheetId)),
      (table = l.tables.find((table) => table.id === tableId)),
      request.kind === "appendRows" &&
        !s.columns.some((column) => column.calculatedFormula))
    ) {
      for (
        let firstRow = rangeAddress.firstRow;
        firstRow <= g.lastRow;
        firstRow++
      )
        for (let firstCol = g.firstCol; firstCol <= g.lastCol; firstCol++) {
          const I = Cn(firstRow, firstCol);
          sheetOrThrow.cells[I] &&
            (l.cells[I] = structuredClone(sheetOrThrow.cells[I]));
        }
      s.showTotals && moveTotalsRowCells(workbook, next, sheetId, s, table);
    }
    if (b)
      for (let k = 0; k < b.length; k++)
        for (let F = 0; F < s.columns.length; F++) {
          const raw = b[k][F];
          if (s.columns[F].calculatedFormula) {
            raw !== null &&
              raw !== "" &&
              throwWorkbookError(
                "Leave calculated-column input slots blank; their saved formulas supply appended values.",
              );
            continue;
          }
          coerceImportedValue(raw, s.columns[F].dataType) ||
            throwWorkbookError(
              `Appended value in ${s.columns[F].name} does not match its ${s.columns[F].dataType} type. Supply a typed value or explicitly revise the column type; no rows were appended.`,
            );
          const O = Cn(g.lastRow + 1 + k, g.firstCol + F),
            cellRecordInput = normalizeCellRecordInput({
              ...l.cells[O],
              raw: raw,
              ...(typeof raw == "string" && raw.startsWith("=")
                ? { literal: true }
                : { literal: undefined }),
            });
          l.cells[O] = cellRecordInput;
        }
  } else throwWorkbookError("Unknown calculated-table operation.");
  (normalizeTableDefinitions(l, l), assertWithinCellBudget(next));
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
      const y = sheetOrThrow.cells[address],
        w = sheet.cells[address];
      (JSON.stringify(y) !== JSON.stringify(w) &&
        changes.push({
          sheetId: sheet.id,
          address: address,
          ...(y ? { before: structuredClone(y) } : {}),
          ...(w ? { after: structuredClone(w) } : {}),
        }),
        changes.length > tableMutationLimits.changedCells &&
          throwWorkbookError(
            "A reviewed table operation may change at most 500 cells.",
          ));
    }
  }
  checkDataValidationBatch(next, changes);
  const anchorChanges = [];
  for (const [v, g, b] of [
    ["watch", workbook.formulaWatches, next.formulaWatches],
    ["review", workbook.reviewThreads, next.reviewThreads],
  ])
    for (const y of b ?? []) {
      const w = g?.find((E) => E.id === y.id);
      w &&
        w.anchor.address !== y.anchor.address &&
        anchorChanges.push({
          kind: v,
          id: y.id,
          before: w.anchor.address,
          after: y.anchor.address,
        });
    }
  return {
    snapshot: JSON.stringify(workbook),
    sheetId: sheetId,
    tableId: tableId,
    request: structuredClone(request),
    beforeRange: s.range,
    afterRange: table.range,
    beforeColumns: structuredClone(s.columns),
    afterColumns: structuredClone(table.columns),
    changes: changes,
    anchorChanges: anchorChanges,
    next: next,
  };
}
function applyTableOperationPreview(workbook, plan, t = {}) {
  plan.snapshot !== JSON.stringify(workbook) &&
    throwWorkbookError(
      "The workbook changed after the table preview. Preview again.",
    );
  const n = applyTableOperation(
    workbook,
    plan.sheetId,
    plan.tableId,
    plan.request,
    t,
  );
  return (
    JSON.stringify(n) !== JSON.stringify(plan) &&
      throwWorkbookError("The table preview was changed. Preview again."),
    JSON.stringify(n.next) === JSON.stringify(workbook) ? workbook : n.next
  );
}
function normalizeTextCleanups(textCleanups, r) {
  if (textCleanups === undefined) return;
  (!Array.isArray(textCleanups) ||
    textCleanups.length > 100 ||
    JSON.stringify(textCleanups).length > 5e6) &&
    throwTextCleanupEvidenceError();
  const t = new Set();
  let n = 0;
  r.textCleanups = textCleanups.map((textCleanup) => {
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
    for (const l of ["id", "workbookId", "sourceSheetId", "outputSheetId"])
      jm(textCleanup[l], 100) || throwTextCleanupEvidenceError();
    ((!jm(textCleanup.sourceSheetName, 31) ||
      !jm(textCleanup.outputSheetName, 31) ||
      t.has(textCleanup.id) ||
      !Number.isSafeInteger(textCleanup.sourceRevision) ||
      Number(textCleanup.sourceRevision) < 0 ||
      textCleanup.semantics !== "unicode-v2-en-US-explicit-separators" ||
      !Array.isArray(textCleanup.rows) ||
      !textCleanup.rows.length ||
      textCleanup.rows.length > 500 ||
      (n += textCleanup.rows.length) > 1e4 ||
      typeof textCleanup.createdAt != "string" ||
      !/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}Z$/.test(
        textCleanup.createdAt,
      ) ||
      !Number.isFinite(Date.parse(textCleanup.createdAt)) ||
      new Date(textCleanup.createdAt).toISOString() !==
        textCleanup.createdAt) &&
      throwTextCleanupEvidenceError(),
      t.add(textCleanup.id));
    const steps = validateTextCleanupSteps(textCleanup.steps),
      s = new Set(),
      d = new Set(),
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
            s.has(row.sourceAddress) ||
            d.has(row.outputAddress) ||
            !WU(row.result)) &&
            throwTextCleanupEvidenceError(),
          s.add(row.sourceAddress),
          d.add(row.outputAddress));
        const original =
          row.original === undefined
            ? undefined
            : normalizeCellRecordInput(row.original);
        original &&
          typeof original.raw == "string" &&
          original.raw.startsWith("=") &&
          !original.literal &&
          throwTextCleanupEvidenceError();
        let h;
        if (row.origin !== undefined) {
          assertPlainObjectKeys(row.origin, [
            "id",
            "sourceId",
            "sourceName",
            "sourceSheet",
            "originalRow",
            "kind",
          ]);
          const f = row.origin;
          ((!jm(f.id, 100) ||
            !jm(f.sourceId, 100) ||
            !jm(f.sourceName, 255) ||
            !jm(f.sourceSheet, 31) ||
            !Number.isInteger(f.originalRow) ||
            Number(f.originalRow) < 1 ||
            Number(f.originalRow) > 1048576 ||
            !["snapshot", "file"].includes(String(f.kind))) &&
            throwTextCleanupEvidenceError(),
            (h = { ...f }));
        }
        return {
          sourceAddress: row.sourceAddress,
          outputAddress: row.outputAddress,
          result: row.result,
          ...(original ? { original: original } : {}),
          ...(h ? { origin: h } : {}),
        };
      });
    return { ...textCleanup, steps: steps, rows: rows };
  });
}
function validateFieldTransformEvidence(fieldTransforms, r) {
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
  const t = new Set();
  let n = 0;
  r.fieldTransforms = fieldTransforms.map((fieldTransform) => {
    const a = analysisOutputEvidenceSchema.safeParse(fieldTransform);
    if (!a.success)
      throw Error(
        `Invalid field transformation evidence: ${a.error.issues[0]?.message}`,
      );
    const data = a.data,
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
    if (t.has(data.id) || (n += data.outputs.length) > 2e4)
      throw Error("Duplicate or over-capacity field transformation evidence.");
    if (
      (t.add(data.id),
      data.sourceSheetId !== request.sourceSheetId ||
        data.outputSheetName !== request.outputName)
    )
      throw Error("Field transformation evidence does not match its recipe.");
    const l = new Set(),
      sources = data.sources.map((source) => {
        const m = _t(source.address);
        if (
          m.row < rangeAddress.firstRow ||
          m.row > rangeAddress.lastRow ||
          m.col < rangeAddress.firstCol ||
          m.col > rangeAddress.lastCol
        )
          throw Error(
            "An original field source is outside the recorded range.",
          );
        if (l.has(source.address))
          throw Error("Duplicate original field source.");
        l.add(source.address);
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
      h = new Set();
    for (const output of data.outputs) {
      if (h.has(output.address))
        throw Error("Duplicate transformed output address.");
      if (
        (h.add(output.address),
        ((output.type === "number" || output.type === "date") &&
          typeof output.value != "number") ||
          (output.type === "boolean" && typeof output.value != "boolean") ||
          (output.type === "text" && typeof output.value != "string") ||
          (output.type === "blank" && output.value !== null))
      )
        throw Error("Transformed output type does not match its value.");
    }
    for (const p of [...data.outputs, ...data.exceptions])
      if (
        p.row < rangeAddress.firstRow + 1 + Number(request.header) ||
        p.row > rangeAddress.lastRow + 1 ||
        p.sources.some((source) => !l.has(source)) ||
        new Set(p.sources).size !== p.sources.length
      )
        throw Error("Unknown or repeated original field source locator.");
    return { ...data, request: request, sources: sources };
  });
}
function normalizeWorkbookModel(workbook, r) {
  if (r && isDeepFrozenValue(workbook) && NC.has(workbook)) return workbook;
  if (!isNonArrayObject(workbook) || workbook.schemaVersion !== 1)
    return throwWithMessage(
      "Unsupported workbook schema. Expected schemaVersion 1.",
    );
  const t = requireBoundedText(workbook.id, "Workbook ID", 100),
    n = requireBoundedText(workbook.name, "Workbook name", 200);
  if (!t || !n.trim())
    return throwWithMessage("Workbook ID and name are required.");
  const revision = P2(
    workbook.revision,
    0,
    Number.MAX_SAFE_INTEGER - 1,
    "Revision",
  );
  if (!Array.isArray(workbook.sheets) || !workbook.sheets.length)
    return throwWithMessage("A workbook must contain at least one sheet.");
  let a = 0;
  const s = new Set(),
    d = new Set(),
    sheets = workbook.sheets.map((sheet) => {
      if (!isNonArrayObject(sheet)) return throwWithMessage("Invalid sheet.");
      const m = requireBoundedText(sheet.id, "Sheet ID", 100),
        v = requireBoundedText(sheet.name, "Sheet name", 31);
      if (
        !m ||
        !v.trim() ||
        /[\\/\[\]*?:]/.test(v) ||
        s.has(m) ||
        d.has(v.toLowerCase())
      )
        return throwWithMessage(
          "Sheet IDs and names must be nonempty and unique; sheet names cannot contain Excel-reserved characters.",
        );
      (s.add(m), d.add(v.toLowerCase()));
      const rows = P2(sheet.rowCount, 1, workbookLimits.maxRows, "Row count"),
        cols = P2(sheet.colCount, 1, workbookLimits.maxCols, "Column count");
      if (!isNonArrayObject(sheet.cells))
        return throwWithMessage("Sheet cells must be an address map.");
      const y = r && isDeepFrozenValue(sheet) ? YE.get(sheet) : undefined;
      if (y !== undefined) {
        if (((a += y), a > workbookLimits.maxCells))
          return throwWithMessage(
            "Workbook exceeds the 100,000 populated cell limit.",
          );
        const A = { ...sheet };
        return (YE.set(A, y), A);
      }
      const cells = {};
      for (const [A, cell] of Object.entries(sheet.cells)) {
        const F = _t(A);
        if (
          A !== `${columnIndexToLetters(F.col)}${F.row + 1}` ||
          F.row >= rows ||
          F.col >= cols
        )
          return throwWithMessage(`Invalid or out-of-bounds cell: ${A}`);
        if (++a > workbookLimits.maxCells)
          return throwWithMessage(
            "Workbook exceeds the 100,000 populated cell limit.",
          );
        cells[A] = normalizeCellRecord(cell, r);
      }
      const E = {
        id: m,
        name: v,
        rowCount: workbookLimits.maxRows,
        colCount: workbookLimits.maxCols,
        cells: cells,
      };
      if (sheet.columnWidths !== undefined) {
        if (!isNonArrayObject(sheet.columnWidths))
          return throwWithMessage("Column widths must be an index map.");
        E.columnWidths = {};
        for (const [A, k] of Object.entries(sheet.columnWidths)) {
          if (
            !/^(0|[1-9]\d*)$/.test(A) ||
            Number(A) >= cols ||
            typeof k != "number" ||
            !Number.isFinite(k) ||
            k < minColumnWidth ||
            k > maxColumnWidth
          )
            return throwWithMessage(
              "Column widths must be 40–600 pixels for an existing column.",
            );
          E.columnWidths[A] = k;
        }
      }
      if (sheet.rowHeights !== undefined) {
        if (
          !isNonArrayObject(sheet.rowHeights) ||
          (Object.getPrototypeOf(sheet.rowHeights) !== Object.prototype &&
            Object.getPrototypeOf(sheet.rowHeights) !== null)
        )
          return throwWithMessage("Row heights must be a plain index map.");
        E.rowHeights = {};
        for (const [A, k] of Object.entries(sheet.rowHeights)) {
          if (
            !/^(0|[1-9]\d*)$/.test(A) ||
            Number(A) >= rows ||
            typeof k != "number" ||
            !Number.isFinite(k) ||
            k < minRowHeight ||
            k > maxRowHeight
          )
            return throwWithMessage(
              `Row heights must be ${minRowHeight}–${maxRowHeight} pixels for an existing row.`,
            );
          E.rowHeights[A] = k;
        }
      }
      if (sheet.hidden !== undefined) {
        if (typeof sheet.hidden != "boolean")
          return throwWithMessage("Sheet hidden flag must be boolean.");
        E.hidden = sheet.hidden;
      }
      return (
        validateWorksheetFeatures(sheet, E),
        validateOutlineGroups(sheet, E, { rows: rows, cols: cols }),
        validateFreezePanes(sheet, E, { rows: rows, cols: cols }),
        validateWorksheetVisibility(sheet, E),
        validateDelimitedSourceFootprint(sheet.delimitedSource, E),
        normalizeTableDefinitions(sheet, E),
        sheet.objects !== undefined && validateDrawingObjects(sheet.objects, E),
        sheet.pageLayout !== undefined &&
          validatePageSetup(sheet.pageLayout, E),
        E
      );
    });
  if (sheets.every((sheet) => sheet.hidden))
    return throwWithMessage("At least one sheet must be visible.");
  if (!Array.isArray(workbook.events) || workbook.events.length > 2e3)
    return throwWithMessage(
      "Workbook history must be an array of at most 2,000 events.",
    );
  const l = new Set();
  let c = 0;
  const events = workbook.events.map((p) => {
      if (!isNonArrayObject(p))
        return throwWithMessage("Invalid history event.");
      const m = requireBoundedText(p.id, "Event ID", 100),
        timestamp = requireBoundedText(p.timestamp, "Event timestamp", 40);
      return !m || l.has(m) || !Number.isFinite(Date.parse(timestamp))
        ? throwWithMessage(
            "History events require unique IDs and valid timestamps.",
          )
        : (l.add(m),
          !Array.isArray(p.changes) ||
          p.changes.length > workbookLimits.maxPatches ||
          (c += p.changes.length) > workbookLimits.maxCells
            ? throwWithMessage("Invalid history change count.")
            : {
                id: m,
                timestamp: timestamp,
                actor: requireBoundedText(p.actor, "Actor", 100),
                action: requireBoundedText(p.action, "Action", 200),
                changes: p.changes.map((change) => {
                  if (
                    !isNonArrayObject(change) ||
                    typeof change.sheetId != "string" ||
                    !change.sheetId.length ||
                    change.sheetId.length > 100 ||
                    typeof change.address != "string"
                  )
                    return throwWithMessage("Invalid history cell reference.");
                  const b = _t(change.address);
                  return change.address !==
                    `${columnIndexToLetters(b.col)}${b.row + 1}`
                    ? throwWithMessage("History cell address is not canonical.")
                    : change.before === undefined && change.after === undefined
                      ? throwWithMessage(
                          "History change must have a before or after value.",
                        )
                      : {
                          sheetId: change.sheetId,
                          address: change.address,
                          ...(change.before !== undefined
                            ? { before: normalizeCellRecord(change.before, r) }
                            : {}),
                          ...(change.after !== undefined
                            ? { after: normalizeCellRecord(change.after, r) }
                            : {}),
                        };
                }),
              });
    }),
    f = {
      schemaVersion: 1,
      id: t,
      name: n,
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
    buildSavedReportEnvelope(workbook, f),
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
    reconcileWorkbookStyleBindings(workbook.styleCatalog, f),
    workbook.formulaWatches !== undefined &&
      (f.formulaWatches = validateFormulaWatches(workbook.formulaWatches, f)),
    zie(workbook, f),
    copyExtensionCacheEntries(workbook, f),
    f
  );
}
function normalizeWorkbookInput(workbook) {
  return normalizeWorkbookModel(workbook, false);
}
function assertThemePreviewCurrent(workbook, plan, t = {}) {
  (t.readOnly && failWithMessage("This workbook is read-only."),
    (plan.version !== 1 || plan.snapshot !== JSON.stringify(workbook)) &&
      failWithMessage(
        "The workbook changed after this preview. Close and reopen Themes & Cell Styles.",
      ));
  const n = applyNamedStyleChange(workbook, plan.request, t);
  if (
    (canonicalStyleKey(n) !== canonicalStyleKey(plan) &&
      failWithMessage(
        "This theme preview changed. Create a fresh preview before applying it.",
      ),
    !plan.count &&
      canonicalStyleKey(plan.catalogBefore) ===
        canonicalStyleKey(plan.catalogAfter))
  )
    return workbook;
  const o = cloneStyleValue(workbook);
  plan.catalogAfter
    ? (o.styleCatalog = cloneStyleValue(plan.catalogAfter))
    : delete o.styleCatalog;
  for (const change of plan.changes)
    findSheetOrThrow(o, change.sheetId).cells[change.address] = cloneStyleValue(
      change.after,
    );
  return normalizeWorkbookInput(o);
}
function Mw(workbook) {
  const workbookInput = normalizeWorkbookInput(workbook);
  return (
    validateWorkbookMetadata(workbook.metadata, workbookInput),
    workbookInput
  );
}
function updateWorkbookProperties(workbook, r) {
  if (
    !isPlainObjectValue(r) ||
    Object.keys(r).some((n) => !["name", "engagement", "period"].includes(n))
  )
    throw new Error("Invalid workbook property fields.");
  const t = Mw(workbook);
  if (
    (r.name !== undefined &&
      (t.name = trimRequiredMetadataText(r.name, "Workbook name", 200)),
    r.engagement !== undefined || r.period !== undefined)
  ) {
    if (typeof r.engagement == "string" && r.engagement.length > 200)
      throw new Error("Engagement must contain at most 200 characters.");
    if (typeof r.period == "string" && r.period.length > 100)
      throw new Error("Period must contain at most 100 characters.");
    const metadata = { ...(t.metadata ?? createWorkbookMetadata()) };
    (r.engagement !== undefined &&
      (typeof r.engagement == "string" && !r.engagement.trim()
        ? delete metadata.engagement
        : (metadata.engagement = trimRequiredMetadataText(
            r.engagement,
            "Engagement",
            200,
          ))),
      r.period !== undefined &&
        (typeof r.period == "string" && !r.period.trim()
          ? delete metadata.period
          : (metadata.period = trimRequiredMetadataText(
              r.period,
              "Period",
              100,
            ))),
      validateWorkbookMetadata(metadata, t));
  }
  return t;
}
function renameWorkbook(activeWorkbook, r) {
  return updateWorkbookProperties(activeWorkbook, { name: r });
}
function cloneWorkbookWithNewIds(e) {
  const workbook = structuredClone(Mw(e)),
    t = new Map(
      workbook.sheets.map((sheet) => [sheet.id, crypto.randomUUID()]),
    ),
    n = new Map(
      workbook.sheets.flatMap((sheet) =>
        (sheet.tables ?? []).map((s) => [s.id, crypto.randomUUID()]),
      ),
    );
  ((workbook.id = crypto.randomUUID()),
    (workbook.revision = 0),
    (workbook.events = []));
  for (const sheet of workbook.sheets)
    if (sheet.analysisReport) {
      const request = sheet.analysisReport.request;
      if (request.kind === "statistics")
        request.sheetId = t.get(request.sheetId) ?? request.sheetId;
      else
        for (const d of [
          request.objective,
          ...request.variables.map((variable) => variable.cell),
          ...request.constraints.flatMap((constraint) => [
            constraint.left,
            ...(typeof constraint.right == "number" ? [] : [constraint.right]),
          ]),
        ])
          d.sheetId = t.get(d.sheetId) ?? d.sheetId;
      sheet.analysisReport.definitionHash = fingerprintCanonicalValue(request);
    }
  const o = [
    workbook.dataModel,
    ...workbook.sheets.map((sheet) => sheet.dataModelReport?.definition),
  ].filter(Boolean);
  for (const a of o)
    for (const table of a.tables)
      ((table.sheetId = t.get(table.sheetId) ?? table.sheetId),
        (table.tableId &&= n.get(table.tableId) ?? table.tableId));
  for (const sheet of workbook.sheets)
    sheet.dataModelReport &&
      (sheet.dataModelReport.definitionHash = fingerprintCanonicalValue(
        sheet.dataModelReport.definition,
      ));
  for (const a of workbook.savedFilterViews ?? [])
    ((a.id = crypto.randomUUID()), (a.sheetId = t.get(a.sheetId) ?? a.sheetId));
  for (const a of workbook.transformationRecipes ?? []) {
    a.id = crypto.randomUUID();
    for (const version of a.versions) {
      for (const d of [
        version.definition.source,
        ...version.definition.steps.flatMap((step) =>
          step.kind === "append" || step.kind === "merge" ? [step.source] : [],
        ),
      ])
        ((d.sheetId = t.get(d.sheetId) ?? d.sheetId),
          (d.tableId &&= n.get(d.tableId) ?? d.tableId));
      version.definitionHash = fingerprintCanonicalValue(version.definition);
    }
  }
  for (const sheet of workbook.sheets) {
    ((sheet.id = t.get(sheet.id)),
      (sheet.outlineGroups &&= regenerateOutlineGroupIds(sheet)));
    for (const s of sheet.tables ?? []) {
      s.id = n.get(s.id);
      for (const column of s.columns) column.id = crypto.randomUUID();
    }
    for (const s of [...(sheet.pivots ?? []), ...(sheet.queries ?? [])])
      ((s.id = crypto.randomUUID()),
        (s.source.sheetId = t.get(s.source.sheetId) ?? s.source.sheetId),
        (s.source.tableId &&= n.get(s.source.tableId) ?? s.source.tableId),
        (s.sourceRevision = 0));
    for (const s of [
      ...(sheet.objects ?? []),
      ...(sheet.validations ?? []),
      ...(sheet.conditionalFormats ?? []),
    ])
      s.id = crypto.randomUUID();
  }
  for (const a of workbook.definedNames ?? [])
    a.scopeSheetId &&= t.get(a.scopeSheetId) ?? a.scopeSheetId;
  return Mw(
    createReviewThread(
      createFormulaWatch(
        renameStyleNamespace(workbook, crypto.randomUUID()),
        t,
      ),
      t,
    ),
  );
}
function createWorkbookFromTemplate(workbook, r) {
  const t = trimRequiredMetadataText(r.name, "Workbook name", 200),
    engagement = trimRequiredMetadataText(r.engagement, "Engagement", 200),
    period = trimRequiredMetadataText(r.period, "Period", 100),
    a = Mw(workbook),
    s = cloneWorkbookWithNewIds(a);
  return (
    (s.name = t),
    (s.metadata = createWorkbookMetadata({
      engagement: engagement,
      period: period,
      template: {
        id: a.id,
        name: a.name,
        version: a.metadata?.template?.version ?? 1,
        sourceWorkbookId: a.id,
        sourceRevision: a.revision,
      },
    })),
    Mw(s)
  );
}
function createBlankWorkbook(e = "Book1") {
  return normalizeWorkbookInput({
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
function assertConditionalFormatPreviewCurrent(workbook, r, t = {}) {
  if (
    r.workbookId !== workbook.id ||
    r.revision !== workbook.revision ||
    r.snapshotGuard !== JSON.stringify(workbook)
  )
    throw Error(
      "Conditional format preview is stale. Reopen the manager to refresh it.",
    );
  const n = updateConditionalFormatRules(workbook, r.sheetId, r.rules, t);
  if (JSON.stringify(n.changes) !== JSON.stringify(r.changes))
    throw Error(
      "Conditional format preview was modified. Preview the rules again.",
    );
  return n.changes.length
    ? normalizeWorkbookInput({
        ...workbook,
        sheets: workbook.sheets.map((sheet) =>
          sheet.id === r.sheetId
            ? { ...sheet, conditionalFormats: n.rules }
            : sheet,
        ),
      })
    : workbook;
}
function w0(workbook) {
  const workbookInput = normalizeWorkbookInput(aQ(workbook));
  return (
    workbookInput.sheets.forEach((sheet, n) =>
      validateWorksheetVisibility(workbook.sheets[n], sheet),
    ),
    validateWorkbookStructureProtection(
      workbook.structureProtection,
      workbookInput,
    ),
    workbookInput
  );
}
function addWorksheet(activeWorkbook, r) {
  return (
    assertWorkbookStructureEditable(activeWorkbook),
    w0({
      ...activeWorkbook,
      sheets: [
        ...activeWorkbook.sheets,
        createWorksheetRecord(assertUniqueWorksheetName(activeWorkbook, r)),
      ],
    })
  );
}
function renameWorksheet(e, r, t) {
  assertWorkbookStructureEditable(e);
  const worksheetById = findWorksheetById(e, r),
    o = assertUniqueWorksheetName(e, t, r);
  if (worksheetById.name === o) return e;
  const workbook = structuredClone(e);
  for (const sheet of workbook.sheets)
    for (const cell of Object.values(sheet.cells))
      !cell.literal &&
        typeof cell.raw == "string" &&
        cell.raw.startsWith("=") &&
        (cell.raw = rewriteStructuredTableReference(
          cell.raw,
          e,
          worksheetById,
          o,
        ));
  for (const sheet of workbook.sheets)
    for (const d of [
      ...(sheet.conditionalFormats ?? []),
      ...(sheet.validations ?? []),
    ])
      if ((d.kind === "formula" || d.kind === "custom") && d.formula) {
        const u = rewriteStructuredTableReference(
          d.formula,
          e,
          worksheetById,
          o,
        );
        (u !== d.formula && assertSheetUnprotected(sheet), (d.formula = u));
      }
  for (const s of workbook.definedNames ?? [])
    s.refersTo = rewriteStructuredTableReference(
      s.refersTo,
      e,
      worksheetById,
      o,
    );
  for (const sheet of workbook.sheets)
    for (const d of sheet.tables ?? [])
      for (const column of d.columns)
        if (column.calculatedFormula) {
          const l = rewriteStructuredTableReference(
            column.calculatedFormula,
            e,
            worksheetById,
            o,
          );
          (l !== column.calculatedFormula && assertSheetUnprotected(sheet),
            (column.calculatedFormula = l));
        }
  return ((findWorksheetById(workbook, r).name = o), w0(workbook));
}
function moveWorksheet(workbook, r, t) {
  assertWorkbookStructureEditable(workbook);
  const worksheetById = findWorksheetById(workbook, r),
    o = workbook.sheets.indexOf(worksheetById);
  if (
    ((!Number.isInteger(t) || t < 0 || t >= workbook.sheets.length) &&
      throwWorkbookError("Choose a worksheet position inside the workbook."),
    o === t)
  )
    return workbook;
  visitTableFormulasOutsideSheet(workbook, (s) => {
    parseWorksheetQualifiedReference(s).some((workbook) => workbook.sheets) &&
      throwWorkbookError(
        "Reordering worksheets with three-dimensional references requires explicit references first.",
      );
  });
  const sheets = [...workbook.sheets];
  return (
    sheets.splice(o, 1),
    sheets.splice(t, 0, worksheetById),
    w0({ ...workbook, sheets: sheets })
  );
}
function setWorksheetVisibility(workbook, r, t) {
  assertWorkbookStructureEditable(workbook);
  const worksheetById = findWorksheetById(workbook, r);
  (["visible", "hidden", "veryHidden"].includes(t) ||
    throwWorkbookError("Choose visible, hidden or very hidden."),
    t !== "visible" &&
      isVisibleWorksheet(worksheetById) &&
      workbook.sheets.filter(isVisibleWorksheet).length === 1 &&
      throwWorkbookError("At least one worksheet must remain visible."));
  const o = worksheetById.veryHidden
    ? "veryHidden"
    : worksheetById.hidden
      ? "hidden"
      : "visible";
  return t === o
    ? workbook
    : w0({
        ...workbook,
        sheets: workbook.sheets.map((sheet) =>
          sheet.id === r
            ? {
                ...sheet,
                hidden: t !== "visible",
                veryHidden: t === "veryHidden",
              }
            : sheet,
        ),
      });
}
function setWorksheetTabColor(workbook, r, tabColor) {
  assertWorkbookStructureEditable(workbook);
  const worksheetById = findWorksheetById(workbook, r),
    o = tabColor === undefined ? undefined : validateTabColor(tabColor);
  return worksheetById.tabColor === o
    ? workbook
    : w0({
        ...workbook,
        sheets: workbook.sheets.map((sheet) => {
          if (sheet.id !== r) return sheet;
          const s = { ...sheet };
          return (o ? (s.tabColor = o) : delete s.tabColor, s);
        }),
      });
}
function duplicateWorksheet(workbook, r, t) {
  assertWorkbookStructureEditable(workbook);
  const worksheetById = findWorksheetById(workbook, r),
    o = assertUniqueWorksheetName(
      workbook,
      createUniqueWorksheetName(workbook, worksheetById.name),
    );
  (visitTableFormulasOutsideSheet(workbook, (g) => {
    parseWorksheetQualifiedReference(g).some((workbook) => workbook.sheets) &&
      throwWorkbookError(
        "Duplicating worksheets with three-dimensional references requires explicit references first.",
      );
  }),
    workbook.sheets.reduce((g, b) => g + Object.keys(b.cells).length, 0) +
      Object.keys(worksheetById.cells).length >
      workbookLimits.maxCells &&
      throwWorkbookError(
        "Duplicating this worksheet would exceed the 100,000 populated-cell budget.",
      ));
  const sheet = structuredClone(worksheetById);
  ((sheet.id = createUuid()),
    (sheet.outlineGroups &&= regenerateOutlineGroupIds(sheet)),
    (sheet.name = o));
  const d = new Set([
      ...workbook.sheets.flatMap((sheet) =>
        (sheet.tables ?? []).map((b) => b.name.toLowerCase()),
      ),
      ...(workbook.definedNames ?? []).map((g) => g.name.toLowerCase()),
    ]),
    u = new Map(),
    l = new Map();
  for (const g of sheet.tables ?? []) {
    const b = g.id,
      y = g.name;
    ((g.id = createUuid()),
      (g.columns = g.columns.map((column) => ({
        ...column,
        id: createUuid(),
      }))));
    for (let w = 2; ; w++) {
      const E = `_Copy${w}`,
        A = y.slice(0, 255 - E.length) + E;
      if (!d.has(A.toLowerCase())) {
        ((g.name = A), d.add(A.toLowerCase()));
        break;
      }
    }
    (u.set(b, g.id), l.set(y, g.name));
  }
  const c = (g) => {
    let b = rewriteStructuredTableReference(g, workbook, worksheetById, o);
    for (const [y, w] of l) b = $v(b, y, w);
    return b;
  };
  for (const g of sheet.tables ?? [])
    for (const column of g.columns)
      column.calculatedFormula &&= c(column.calculatedFormula);
  for (const cell of Object.values(sheet.cells))
    !cell.literal &&
      typeof cell.raw == "string" &&
      cell.raw.startsWith("=") &&
      (cell.raw = c(cell.raw));
  (sheet.objects?.forEach((object) => {
    object.id = createUuid();
  }),
    sheet.validations?.forEach((validation) => {
      ((validation.id = createUuid()),
        validation.kind === "custom" &&
          validation.formula &&
          (validation.formula = c(validation.formula)));
    }),
    sheet.conditionalFormats?.forEach((conditionalFormat) => {
      ((conditionalFormat.id = createUuid()),
        conditionalFormat.kind === "formula" &&
          conditionalFormat.formula &&
          (conditionalFormat.formula = c(conditionalFormat.formula)));
    }));
  for (const g of [...(sheet.pivots ?? []), ...(sheet.queries ?? [])])
    ((g.id = createUuid()),
      (g.sourceRevision = workbook.revision),
      g.source.sheetId === worksheetById.id &&
        ((g.source.sheetId = sheet.id),
        (g.source.tableId &&=
          u.get(g.source.tableId) ??
          throwWorkbookError(
            "A copied analysis refers to a missing source table.",
          ))));
  const h = (workbook.definedNames ?? [])
      .filter((g) => g.scopeSheetId === worksheetById.id)
      .map((g) => ({
        ...structuredClone(g),
        scopeSheetId: sheet.id,
        refersTo: c(g.refersTo),
      })),
    sheets = [...workbook.sheets];
  return (
    sheets.splice(workbook.sheets.indexOf(worksheetById) + 1, 0, sheet),
    w0(
      collectHyperlinkEntries(
        {
          ...workbook,
          sheets: sheets,
          ...(h.length
            ? { definedNames: [...(workbook.definedNames ?? []), ...h] }
            : {}),
        },
        worksheetById.id,
        sheet.id,
      ),
    )
  );
}
function deleteWorksheet(workbook, sheetId, t = false) {
  assertWorkbookStructureEditable(workbook);
  const worksheetById = findWorksheetById(workbook, sheetId);
  (workbook.sheets.length === 1 &&
    throwWorkbookError("The final worksheet cannot be deleted."),
    isVisibleWorksheet(worksheetById) &&
      workbook.sheets.filter(isVisibleWorksheet).length === 1 &&
      throwWorkbookError("At least one worksheet must remain visible."));
  const o = describeSheetDeletionImpacts(workbook, sheetId);
  (o.some((s) => s.kind === "unsupported-reference") &&
    throwWorkbookError(
      "Deletion is blocked by ambiguous or unsupported references. Replace dynamic or three-dimensional references with explicit references first.",
    ),
    o.length &&
      !t &&
      throwWorkbookError(
        "This worksheet has dependencies. Review them and explicitly confirm deletion.",
      ));
  const a = structuredClone(workbook);
  a.sheets = a.sheets.filter((sheet) => sheet.id !== sheetId);
  for (const sheet of a.sheets) {
    for (const d of sheet.tables ?? [])
      for (const column of d.columns)
        if (column.calculatedFormula) {
          const l = rewriteStructuredTableReference(
            column.calculatedFormula,
            workbook,
            worksheetById,
            undefined,
            true,
          );
          (l !== column.calculatedFormula && assertSheetUnprotected(sheet),
            (column.calculatedFormula = l));
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
    for (const d of [
      ...(sheet.conditionalFormats ?? []),
      ...(sheet.validations ?? []),
    ])
      if ((d.kind === "formula" || d.kind === "custom") && d.formula) {
        const u = rewriteStructuredTableReference(
          d.formula,
          workbook,
          worksheetById,
          undefined,
          true,
        );
        (u !== d.formula && assertSheetUnprotected(sheet), (d.formula = u));
      }
    ((sheet.pivots &&= sheet.pivots.filter(
      (pivot) => pivot.source.sheetId !== sheetId,
    )),
      (sheet.queries &&= sheet.queries.filter(
        (query) => query.source.sheetId !== sheetId,
      )));
  }
  return (
    (a.definedNames &&= a.definedNames
      .filter((definedName) => definedName.scopeSheetId !== sheetId)
      .map((s) => ({
        ...s,
        refersTo: rewriteStructuredTableReference(
          s.refersTo,
          workbook,
          worksheetById,
          undefined,
          true,
        ),
      }))),
    w0(fQ(a, sheetId))
  );
}
async function protectWorkbookStructure(workbook, r) {
  (assertWorkbookStructureEditable(workbook), validateWorkbookPassword(r));
  const uuid = createUuid(),
    structureProtection = r
      ? {
          enabled: true,
          salt: uuid,
          passwordHash: await hashWorkbookStructurePassword(r, uuid),
        }
      : { enabled: true };
  return w0({ ...workbook, structureProtection: structureProtection });
}
async function unprotectWorkbookStructure(workbook, r) {
  validateWorkbookPassword(r);
  const structureProtection = workbook.structureProtection;
  if (!structureProtection) return workbook;
  (validateWorkbookStructureProtection(structureProtection, { ...workbook }),
    structureProtection.passwordHash &&
      (await hashWorkbookStructurePassword(r, structureProtection.salt)) !==
        structureProtection.passwordHash.toLowerCase() &&
      throwWorkbookError("The workbook structure password does not match."));
  const n = { ...workbook };
  return (delete n.structureProtection, w0(n));
}
function t$(workbookModel) {
  deepFreezeValue(workbookModel);
  for (const sheet of workbookModel.sheets) {
    if (YE.has(sheet)) continue;
    const t = Object.values(sheet.cells);
    for (const n of t) JU.add(n);
    (QU.set(sheet.cells, t.length),
      t.some((n) => n.hyperlink || n.style?.binding) ||
        YE.set(sheet, t.length));
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
