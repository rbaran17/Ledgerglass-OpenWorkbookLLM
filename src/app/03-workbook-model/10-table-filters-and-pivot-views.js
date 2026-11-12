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

function parseReportSnapshot(snapshot) {
  if (new TextEncoder().encode(JSON.stringify(snapshot)).length > 4e6)
    throw Error("A report snapshot exceeds 4 MB.");
  const report = calculationReportSchema.parse(snapshot),
    rangeAddress = parseRangeAddress(report.request.range);
  if (
    (validateObjectReportRequest(report.request),
    report.hash !== hashRecordExcludingHash(report))
  )
    throw Error("Report snapshot integrity check failed.");
  const columnIndexes = Array.from(
    { length: rangeAddress.lastCol - rangeAddress.firstCol + 1 },
    (s, columnOffset) => rangeAddress.firstCol + columnOffset,
  );
  if (
    fingerprintCanonicalValue(columnIndexes) !==
      fingerprintCanonicalValue(report.columns) ||
    report.widths.length !== columnIndexes.length
  )
    throw Error("Invalid report column layout.");
  const rowNumbers = report.rows.map((row) => row.row),
    allRowNumbers = [...rowNumbers, ...report.excludedRows].sort(
      (left, right) => left - right,
    );
  if (
    report.rows.length * columnIndexes.length > 2e3 ||
    allRowNumbers.length !== rangeAddress.lastRow - rangeAddress.firstRow + 1 ||
    allRowNumbers.some(
      (rowNumber, rowIndex) =>
        rowNumber !== rangeAddress.firstRow + rowIndex + 1,
    ) ||
    rowNumbers.some(
      (candidateRow, rowPosition) =>
        rowPosition > 0 && candidateRow <= rowNumbers[rowPosition - 1],
    )
  )
    throw Error("Invalid report row accounting.");
  for (const row of report.rows) {
    if (row.cells.length !== columnIndexes.length)
      throw Error("Invalid report cell layout.");
    row.cells.forEach((cell, cellIndex) => {
      if (
        cell.address !==
        `${columnIndexToLetters(columnIndexes[cellIndex])}${row.row}`
      )
        throw Error("Invalid report cell address.");
      cell.style &&= validateCellStyle(cell.style);
    });
  }
  if (report.object) {
    const drawingSheet = {
      id: report.request.sheetId,
      name: report.sheetName,
      rowCount: 1048576,
      colCount: 16384,
    };
    (validateDrawingObjects([report.object], drawingSheet),
      (report.object = drawingSheet.objects[0]));
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
function buildSavedReportEnvelope(workbook, envelope) {
  if (workbook.reportOutputs !== undefined) {
    const parsedReportOutputs = array(
      object({
        id: nonEmptyLabelSchema,
        version: number().int().min(1).max(1e6),
        snapshot: unknown(),
      }).strict(),
    )
      .max(20)
      .parse(workbook.reportOutputs);
    if (
      new Set(parsedReportOutputs.map((reportOutput) => reportOutput.id))
        .size !== parsedReportOutputs.length
    )
      throw Error("Duplicate report output IDs.");
    if (
      new TextEncoder().encode(JSON.stringify(parsedReportOutputs)).length >
      reportOutputLimits.metadataBytes
    )
      throw Error("Saved reports exceed 16 MB.");
    envelope.reportOutputs = parsedReportOutputs.map((outputEntry) => ({
      ...outputEntry,
      snapshot: parseReportSnapshot(outputEntry.snapshot),
    }));
  }
  if (workbook.embeddedObjectInventories !== undefined) {
    const parsedInventories = array(importedCacheEvidenceSchema)
      .max(20)
      .parse(workbook.embeddedObjectInventories);
    if (
      new Set(parsedInventories.map((inventory) => inventory.id)).size !==
        parsedInventories.length ||
      JSON.stringify(parsedInventories).length > 2e6
    )
      throw Error("Invalid embedded-object inventory count or size.");
    envelope.embeddedObjectInventories = parsedInventories;
  }
}
function validatePageSetup(pageLayout, sheet) {
  if (pageLayout === undefined) return;
  if (!pageLayout || typeof pageLayout != "object" || Array.isArray(pageLayout))
    throw new Error("Invalid page layout.");
  const layout = pageLayout,
    margins = layout.margins,
    checkBoundedNumber = (value, minimum, maximum) => {
      if (
        typeof value != "number" ||
        !Number.isFinite(value) ||
        value < minimum ||
        value > maximum
      )
        throw new Error("Page setup contains an out-of-range number.");
      return value;
    };
  if (
    !["portrait", "landscape"].includes(String(layout.orientation)) ||
    !["letter", "a4", "legal"].includes(String(layout.paper)) ||
    !margins ||
    typeof margins != "object"
  )
    throw new Error("Invalid paper, orientation, or margins.");
  if (
    layout.pageOrder !== undefined &&
    layout.pageOrder !== "overThenDown" &&
    layout.pageOrder !== "downThenOver"
  )
    throw Error(
      "Unsupported print page order. Choose Over, then down or Down, then over.",
    );
  const pageMargins = Object.fromEntries(
      ["top", "bottom", "left", "right", "header", "footer"].map(
        (marginName) => [
          marginName,
          checkBoundedNumber(margins[marginName], 0, 3),
        ],
      ),
    ),
    paperSize =
      layout.paper === "a4"
        ? [8.2677, 11.6929]
        : layout.paper === "legal"
          ? [8.5, 14]
          : [8.5, 11],
    [pageWidth, pageHeight] =
      layout.orientation === "landscape" ? paperSize.reverse() : paperSize;
  if (
    pageMargins.left + pageMargins.right >= pageWidth - 0.25 ||
    pageMargins.top + pageMargins.bottom >= pageHeight - 0.25
  )
    throw new Error("Margins leave no printable page area.");
  if (typeof layout.printArea != "string")
    throw new Error("Enter a print area.");
  const rangeAddress = parseRangeAddress(layout.printArea);
  if (
    rangeAddress.lastRow >= sheet.rowCount ||
    rangeAddress.lastCol >= sheet.colCount
  )
    throw new Error("Print area exceeds the worksheet.");
  if (
    ["fitToPage", "gridlines", "headings"].some(
      (optionName) => typeof layout[optionName] != "boolean",
    )
  )
    throw new Error("Invalid print option.");
  if (
    ["header", "footer"].some(
      (fieldName) =>
        typeof layout[fieldName] != "string" ||
        String(layout[fieldName]).length > 255,
    )
  )
    throw new Error("Headers and footers must be at most 255 characters.");
  const fitToWidth = checkBoundedNumber(layout.fitToWidth, 0, 10),
    fitToHeight = checkBoundedNumber(layout.fitToHeight, 0, 10),
    scale = checkBoundedNumber(layout.scale, 10, 400);
  if (![fitToWidth, fitToHeight, scale].every(Number.isInteger))
    throw new Error("Page fitting and scale must be whole numbers.");
  const printTitles = parsePrintTitles(layout.repeatRows, "row", sheet),
    columnTitles = parsePrintTitles(layout.repeatColumns, "column", sheet),
    rowBreaks = validatePrintPageBreaks(layout.rowBreaks, sheet.rowCount),
    columnBreaks = validatePrintPageBreaks(layout.columnBreaks, sheet.colCount);
  sheet.pageLayout = {
    orientation: layout.orientation,
    paper: layout.paper,
    margins: pageMargins,
    printArea: layout.printArea,
    fitToPage: layout.fitToPage,
    fitToWidth: fitToWidth,
    fitToHeight: fitToHeight,
    scale: scale,
    gridlines: layout.gridlines,
    headings: layout.headings,
    header: layout.header,
    footer: layout.footer,
    ...(layout.pageOrder !== undefined ? { pageOrder: layout.pageOrder } : {}),
    ...(printTitles
      ? { repeatRows: formatAxisIntervalRef(printTitles, "row") }
      : {}),
    ...(columnTitles
      ? { repeatColumns: formatAxisIntervalRef(columnTitles, "column") }
      : {}),
    ...(rowBreaks ? { rowBreaks: rowBreaks } : {}),
    ...(columnBreaks ? { columnBreaks: columnBreaks } : {}),
  };
}
function zU(sheet, pageLayout) {
  const sheetCopy = { ...sheet };
  (validatePageSetup(pageLayout, sheetCopy),
    (pageLayout = sheetCopy.pageLayout));
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
    printableWidth =
      pageDimensions.width -
      (pageLayout.margins.left + pageLayout.margins.right) * 96,
    printableHeight =
      pageDimensions.height -
      (pageLayout.margins.top + pageLayout.margins.bottom) * 96,
    printTitles = parsePrintTitles(pageLayout.repeatRows, "row", sheet),
    columnTitles = parsePrintTitles(pageLayout.repeatColumns, "column", sheet);
  for (const mergeRange of sheet.merges ?? []) {
    const rangeAddress = parseRangeAddress(mergeRange);
    for (const [titleRange, firstIndex, lastIndex] of [
      [printTitles, rangeAddress.firstRow, rangeAddress.lastRow],
      [columnTitles, rangeAddress.firstCol, rangeAddress.lastCol],
    ])
      if (
        titleRange &&
        firstIndex <= titleRange.end &&
        lastIndex >= titleRange.start &&
        (firstIndex < titleRange.start || lastIndex > titleRange.end)
      )
        throw Error(
          "A merged area crosses a repeated-title boundary. Unmerge it or adjust the print titles first.",
        );
  }
  for (const object of sheet.objects ?? []) {
    const anchorCell = _t(object.anchor),
      drawingLayout = applyChartDrawingLayout(sheet, object);
    for (const [axisTitles, axis, anchorIndex, axisOffset, axisSize] of [
      [
        printTitles,
        "row",
        anchorCell.row,
        drawingLayout.offsetY,
        drawingLayout.height,
      ],
      [
        columnTitles,
        "column",
        anchorCell.col,
        drawingLayout.offsetX,
        drawingLayout.width,
      ],
    ]) {
      if (!axisTitles) continue;
      const axisVisibilityIndex = createAxisVisibilityIndex(
        axis === "row" ? sheet.rowCount : sheet.colCount,
        cn(sheet, axis),
      );
      if (axisVisibilityIndex.isHidden(anchorIndex)) continue;
      const axisSizeAt = (axisIndex) =>
          axis === "row"
            ? rowHeightAt(sheet, axisIndex)
            : columnWidthAt(sheet, axisIndex),
        indexAtOffset = (offset) => {
          let currentIndex = anchorIndex,
            iterationCount = 0;
          for (; offset < 0;) {
            const previousVisibleIndex = axisVisibilityIndex.next(
              currentIndex,
              -1,
            );
            if (previousVisibleIndex === null) return currentIndex;
            if (
              ((currentIndex = previousVisibleIndex),
              (offset += axisSizeAt(currentIndex)),
              ++iterationCount > 1e3)
            )
              throw Error(
                "Drawing print inspection exceeds its geometry budget.",
              );
          }
          for (; offset >= axisSizeAt(currentIndex);) {
            offset -= axisSizeAt(currentIndex);
            const nextVisibleIndex = axisVisibilityIndex.next(currentIndex, 1);
            if (nextVisibleIndex === null) return currentIndex;
            if (((currentIndex = nextVisibleIndex), ++iterationCount > 1e3))
              throw Error(
                "Drawing print inspection exceeds its geometry budget.",
              );
          }
          return currentIndex;
        },
        startIndex = indexAtOffset(axisOffset),
        endIndex = indexAtOffset(axisOffset + axisSize - 1e-4);
      if (
        startIndex <= axisTitles.end &&
        endIndex >= axisTitles.start &&
        (startIndex < axisTitles.start || endIndex > axisTitles.end)
      )
        throw Error(
          "A drawing crosses a repeated-title boundary. Move or resize the drawing, or adjust the print titles first.",
        );
    }
  }
  const listVisibleIndexes = (axisKind, rangeStart, rangeEnd) => {
      const axisVisibilityIndex = createAxisVisibilityIndex(
          axisKind === "row" ? sheet.rowCount : sheet.colCount,
          cn(sheet, axisKind),
        ),
        visibleIndexes = [];
      for (let scanIndex = rangeStart; scanIndex <= rangeEnd; scanIndex++)
        axisVisibilityIndex.isHidden(scanIndex) ||
          visibleIndexes.push(scanIndex);
      return visibleIndexes;
    },
    visibleRows = listVisibleIndexes(
      "row",
      rangeAddress.firstRow,
      rangeAddress.lastRow,
    ),
    visibleColumns = listVisibleIndexes(
      "column",
      rangeAddress.firstCol,
      rangeAddress.lastCol,
    );
  if (!visibleRows.length || !visibleColumns.length)
    throw Error(
      "The print area has no visible cells. Unhide its rows or columns, or choose another print area.",
    );
  const visibleTitleRows = printTitles
      ? listVisibleIndexes("row", printTitles.start, printTitles.end)
      : [],
    visibleTitleColumns = columnTitles
      ? listVisibleIndexes("column", columnTitles.start, columnTitles.end)
      : [],
    paginateAxis = (
      axisIndexes,
      titleIndexes,
      availableSize,
      measureSize,
      breakIndexes,
    ) => {
      const titleSet = new Set(titleIndexes),
        dataIndexes = axisIndexes.filter(
          (entryIndex) => !titleSet.has(entryIndex),
        ),
        titleSize = titleIndexes.reduce(
          (sizeTotal, titleIndex) => sizeTotal + measureSize(titleIndex),
          0,
        );
      if (
        titleSize > availableSize + 1e-7 ||
        (dataIndexes.length && titleSize >= availableSize - 1e-7)
      )
        throw Error(
          "Repeated titles leave no space for worksheet data. Reduce the titles, margins or scale.",
        );
      if (!dataIndexes.length) return [titleIndexes];
      const pageGroups = [];
      let currentGroup = [],
        usedSize = titleSize,
        lastPlacedIndex;
      for (const dataIndex of dataIndexes) {
        const dataSize = measureSize(dataIndex),
          isBreakBefore =
            lastPlacedIndex !== undefined &&
            breakIndexes.some(
              (breakIndex) =>
                breakIndex > lastPlacedIndex && breakIndex <= dataIndex,
            );
        if (
          (currentGroup.length &&
            (usedSize + dataSize > availableSize + 1e-7 || isBreakBefore) &&
            (pageGroups.push([...titleIndexes, ...currentGroup]),
            (currentGroup = []),
            (usedSize = titleSize)),
          titleIndexes.length &&
            usedSize + dataSize > availableSize + 1e-7 &&
            !currentGroup.length)
        )
          throw Error(
            "A data row or column cannot fit beside the repeated titles. Reduce the scale or title dimensions.",
          );
        (currentGroup.push(dataIndex),
          (usedSize += dataSize),
          (lastPlacedIndex = dataIndex));
      }
      return (
        currentGroup.length &&
          pageGroups.push([...titleIndexes, ...currentGroup]),
        pageGroups
      );
    },
    layoutAtScale = (candidateScale) => ({
      rows: paginateAxis(
        visibleRows,
        visibleTitleRows,
        printableHeight / candidateScale - (pageLayout.headings ? 22 : 0),
        (rowIndex) => rowHeightAt(sheet, rowIndex),
        pageLayout.fitToPage ? [] : (pageLayout.rowBreaks ?? []),
      ),
      columns: paginateAxis(
        visibleColumns,
        visibleTitleColumns,
        printableWidth / candidateScale - (pageLayout.headings ? 32 : 0),
        (columnIndex) => columnWidthAt(sheet, columnIndex),
        pageLayout.fitToPage ? [] : (pageLayout.columnBreaks ?? []),
      ),
    });
  let scale = pageLayout.fitToPage ? 1 : pageLayout.scale / 100;
  if (
    pageLayout.fitToPage &&
    (pageLayout.fitToWidth || pageLayout.fitToHeight)
  ) {
    const canFitAtScale = (trialScale) => {
      try {
        const trialLayout = layoutAtScale(trialScale);
        return (
          (!pageLayout.fitToWidth ||
            trialLayout.columns.length <= pageLayout.fitToWidth) &&
          (!pageLayout.fitToHeight ||
            trialLayout.rows.length <= pageLayout.fitToHeight)
        );
      } catch {
        return false;
      }
    };
    if (!canFitAtScale(1)) {
      let lowScale = 0,
        highScale = 1;
      for (let searchStep = 0; searchStep < 40; searchStep++) {
        const midScale = (lowScale + highScale) / 2;
        canFitAtScale(midScale)
          ? (lowScale = midScale)
          : (highScale = midScale);
      }
      scale = lowScale;
    }
  }
  if (!Number.isFinite(scale) || scale <= 0)
    throw Error(
      "The print content cannot fit the selected paper. Reduce the print area or titles.",
    );
  const finalLayout = layoutAtScale(scale);
  if (
    finalLayout.rows.length * finalLayout.columns.length >
    printRenderLimits.pages
  )
    throw Error(
      "This layout exceeds 100 pages. Reduce the print area or fit it to fewer pages.",
    );
  if (
    finalLayout.rows.reduce(
      (rowTotal, pageRows) => rowTotal + pageRows.length,
      0,
    ) *
      finalLayout.columns.reduce(
        (columnTotal, pageColumns) => columnTotal + pageColumns.length,
        0,
      ) >
    printRenderLimits.renderedCells
  )
    throw Error(
      "Repeated print content exceeds 100,000 rendered cells. Reduce titles, pages or the print area.",
    );
  const hasRepeatedTitles = !!(printTitles || columnTitles),
    toPageRange = (rows, columns) => ({
      firstRow: Math.min(...rows),
      lastRow: Math.max(...rows),
      firstCol: Math.min(...columns),
      lastCol: Math.max(...columns),
      ...(hasRepeatedTitles ? { rows: rows, columns: columns } : {}),
    }),
    pages =
      pageLayout.pageOrder === "downThenOver"
        ? finalLayout.columns.flatMap((column) =>
            finalLayout.rows.map((row) => toPageRange(row, column)),
          )
        : finalLayout.rows.flatMap((row) =>
            finalLayout.columns.map((column) => toPageRange(row, column)),
          ),
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
  const normalizedSheet = { ...sheet };
  validatePageSetup(sheet.pageLayout, normalizedSheet);
  const pageLayout = normalizedSheet.pageLayout,
    oddHeader = exportHeaderFooterTemplate(pageLayout.header, workbookInput),
    oddFooter = exportHeaderFooterTemplate(pageLayout.footer, workbookInput),
    hasIdentityTokens = [pageLayout.header, pageLayout.footer].some(
      (templateText) =>
        [...templateText.matchAll(headerFooterTokenPattern)].some(
          (tokenMatch) => !!tokenMatch[1],
        ),
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
    (e.rowBreaks = (pageLayout.rowBreaks ?? []).map((breakRow) => ({
      id: breakRow,
      max: 16383,
      man: 1,
    }))),
    (e.headerFooter = { oddHeader: oddHeader, oddFooter: oddFooter }),
    hasIdentityTokens
      ? [
          "Print workbook/engagement/period/revision fields are exported as literal text from this workbook snapshot. Native JSON retains the live identity tokens; XLSX does not update them after later edits.",
        ]
      : []
  );
}
function findOccupiedRangeBounds(sheet, shouldIncludeFormatting = true) {
  let bounds = null;
  for (const [cellAddress, cell] of Object.entries(sheet.cells)) {
    const hasValue = cell.raw !== null && cell.raw !== "",
      hasFormatting =
        shouldIncludeFormatting &&
        (cell.numberFormat ||
          cell.hyperlink ||
          cell.note ||
          cell.tickmark ||
          cell.numFmt ||
          cell.format ||
          Object.keys(cell.style ?? {}).length);
    if (!hasValue && !hasFormatting) continue;
    const { row, col } = _t(cellAddress);
    bounds
      ? ((bounds.firstRow = Math.min(bounds.firstRow, row)),
        (bounds.lastRow = Math.max(bounds.lastRow, row)),
        (bounds.firstCol = Math.min(bounds.firstCol, col)),
        (bounds.lastCol = Math.max(bounds.lastCol, col)))
      : (bounds = { firstRow: row, lastRow: row, firstCol: col, lastCol: col });
  }
  return bounds;
}
function formatOccupiedRangeAddress(sheet, shouldIncludeFormatting = true) {
  const occupiedRangeBounds = findOccupiedRangeBounds(
    sheet,
    shouldIncludeFormatting,
  );
  if (!occupiedRangeBounds) return "A1";
  const firstAddress = `${columnIndexToLetters(occupiedRangeBounds.firstCol)}${occupiedRangeBounds.firstRow + 1}`,
    lastAddress = `${columnIndexToLetters(occupiedRangeBounds.lastCol)}${occupiedRangeBounds.lastRow + 1}`;
  return firstAddress === lastAddress
    ? firstAddress
    : `${firstAddress}:${lastAddress}`;
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
function assertRangeCellBudget(rangeText) {
  const { firstRow, lastRow, firstCol, lastCol } = parseRangeAddress(rangeText);
  if (
    (lastRow - firstRow + 1) * (lastCol - firstCol + 1) >
    workbookLimits.maxCells
  )
    return throwWithMessage(
      "Range is too large. Select at most 100,000 cells.",
    );
  const cellAddresses = [];
  for (let rowIndex = firstRow; rowIndex <= lastRow; rowIndex++)
    for (let colIndex = firstCol; colIndex <= lastCol; colIndex++)
      cellAddresses.push(`${columnIndexToLetters(colIndex)}${rowIndex + 1}`);
  return cellAddresses;
}
function validateSelectionBounds(sheetOrThrow, activeCellAddress) {
  const selectionAddresses = assertRangeCellBudget(activeCellAddress);
  if (
    selectionAddresses.some((address) => {
      const { row, col } = _t(address);
      return row >= sheetOrThrow.rowCount || col >= sheetOrThrow.colCount;
    })
  )
    throw new Error("Selection exceeds worksheet dimensions.");
  return selectionAddresses;
}
function validateRemoveDuplicatesRequest(
  workbook,
  sheetId,
  range,
  columnIndexes,
  checked = true,
) {
  const sheet = workbook.sheets.find((sheet) => sheet.id === sheetId);
  assertSheetUnprotected(sheet);
  const cellAddresses = assertRangeCellBudget(range);
  cellAddresses.length > 500 &&
    failWithMessage("Select at most 500 cells to remove duplicates.");
  const rangeAddress = parseRangeAddress(range);
  (!columnIndexes.length ||
    columnIndexes.some(
      (columnIndex) =>
        columnIndex < rangeAddress.firstCol ||
        columnIndex > rangeAddress.lastCol,
    )) &&
    failWithMessage("Choose one or more columns inside the selected range.");
  for (const table of sheet.tables ?? []) {
    const tableRangeAddress = parseRangeAddress(table.range);
    doesRangeOverlap(range, table.range) &&
      ((rangeAddress.firstCol > tableRangeAddress.firstCol ||
        rangeAddress.lastCol < tableRangeAddress.lastCol) &&
        failWithMessage(
          "Include all table columns when removing duplicate table rows.",
        ),
      table.showTotals &&
        rangeAddress.firstRow <= tableRangeAddress.lastRow &&
        rangeAddress.lastRow >= tableRangeAddress.lastRow &&
        failWithMessage("Select data rows above the table totals row."),
      !checked &&
        rangeAddress.firstRow <= tableRangeAddress.firstRow &&
        rangeAddress.lastRow >= tableRangeAddress.firstRow &&
        failWithMessage("Keep the table header row selected as headers."));
  }
  (sheet.merges ?? []).some((mergeRange) =>
    doesRangeOverlap(range, mergeRange),
  ) && failWithMessage("Unmerge selected cells before removing duplicates.");
  const workbookWithSettings = calculateWorkbookWithSettings(workbook),
    seenRowKeys = new Set(),
    copiedRows = [];
  let removed = 0;
  for (
    let firstRow = rangeAddress.firstRow;
    firstRow <= rangeAddress.lastRow;
    firstRow++
  ) {
    const rowKey = JSON.stringify(
      columnIndexes.map((keyColumnIndex) =>
        workbookWithSettings(
          sheetId,
          `${columnIndexToLetters(keyColumnIndex)}${firstRow + 1}`,
        ),
      ),
    );
    if (firstRow !== rangeAddress.firstRow || !checked) {
      if (seenRowKeys.has(rowKey)) {
        removed++;
        continue;
      }
      seenRowKeys.add(rowKey);
    }
    copiedRows.push(
      Array.from(
        { length: rangeAddress.lastCol - rangeAddress.firstCol + 1 },
        (g, columnOffset) => {
          const cell = structuredClone(
            sheet.cells[
              `${columnIndexToLetters(rangeAddress.firstCol + columnOffset)}${firstRow + 1}`
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
                rangeAddress.firstRow + copiedRows.length - firstRow,
                0,
              )),
            cell
          );
        },
      ),
    );
  }
  const cells = { ...sheet.cells };
  for (const cellAddress of cellAddresses) {
    const cellPosition = _t(cellAddress);
    cells[cellAddress] = copiedRows[cellPosition.row - rangeAddress.firstRow]?.[
      cellPosition.col - rangeAddress.firstCol
    ] ?? {
      ...cells[cellAddress],
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
function findSheetOrThrow(workbook, sheetId) {
  return (
    workbook.sheets.find((sheet) => sheet.id === sheetId) ??
    throwWithMessage(`Sheet not found: ${sheetId}`)
  );
}
function applyNamedStyleChange(workbook, request, options = {}) {
  options.readOnly && failWithMessage("This workbook is read-only.");
  for (const sheet of workbook.sheets)
    for (const cell of Object.values(sheet.cells))
      assertNamedStyleBindingCurrent(workbook, cell.style);
  const styleCatalog = workbook.styleCatalog,
    workingCatalog = cloneStyleValue(
      styleCatalog ?? createDefaultStyleCatalog(workbook.id),
    );
  let workingRequest = cloneStyleValue(request);
  const changes = [],
    warnings = [],
    findNamedStyle = (targetStyleId) =>
      workingCatalog.styles.find((style) => style.id === targetStyleId) ??
      failWithMessage("This named style no longer exists.");
  function recordCellStyleChange(sheetId, address, nextCell) {
    const sheetOrThrow = findSheetOrThrow(workbook, sheetId),
      currentCell = sheetOrThrow.cells[address];
    if (canonicalStyleKey(currentCell) !== canonicalStyleKey(nextCell)) {
      if (
        (assertSheetUnprotected(sheetOrThrow),
        hasBorderEdgeChanges(currentCell?.style, nextCell.style))
      ) {
        const cellPosition = _t(address);
        sheetOrThrow.merges?.some((merge) => {
          const rangeAddress = parseRangeAddress(merge);
          return (
            cellPosition.row >= rangeAddress.firstRow &&
            cellPosition.row <= rangeAddress.lastRow &&
            cellPosition.col >= rangeAddress.firstCol &&
            cellPosition.col <= rangeAddress.lastCol
          );
        }) &&
          failWithMessage(
            "Unmerge the intersecting cells before changing named-style border edges.",
          );
      }
      (changes.push({
        sheetId: sheetId,
        address: address,
        ...(currentCell ? { before: cloneStyleValue(currentCell) } : {}),
        after: cloneStyleValue(nextCell),
      }),
        changes.length > 500 &&
          failWithMessage(
            "This action changes more than 500 cells. Use smaller selections or fewer bound cells.",
          ));
    }
  }
  function refreshBoundCells(filterStyleId) {
    for (const sheet of workbook.sheets)
      for (const [cellAddress, cellRecord] of Object.entries(sheet.cells))
        cellRecord.style?.binding &&
          (!filterStyleId ||
            cellRecord.style.binding.styleId === filterStyleId) &&
          (assertSheetUnprotected(sheet),
          recordCellStyleChange(sheet.id, cellAddress, {
            ...cellRecord,
            style: {
              ...cellRecord.style,
              binding: resolveNamedStyleBinding(
                workingCatalog,
                cellRecord.style.binding,
              ),
            },
          }));
  }
  switch (workingRequest.kind) {
    case "setTheme":
      ((workingCatalog.theme = validateWorkbookTheme(workingRequest.theme)),
        delete workingCatalog.missingTheme,
        (workingRequest.theme = cloneStyleValue(workingCatalog.theme)),
        refreshBoundCells());
      break;
    case "createStyle": {
      const styleName = validateBoundedStyleText(
          workingRequest.name,
          "Style name",
        ),
        styleId =
          workingRequest.styleId === undefined
            ? `style-${crypto.randomUUID()}`
            : validateBoundedStyleText(workingRequest.styleId, "Style ID", 100),
        definition = validateNamedStyleDefinition(workingRequest.definition);
      (workingCatalog.styles.push({
        id: styleId,
        name: styleName,
        definition: definition,
      }),
        (workingRequest = {
          ...workingRequest,
          name: styleName,
          styleId: styleId,
          definition: definition,
        }),
        refreshBoundCells(styleId));
      break;
    }
    case "updateStyle": {
      const styleToUpdate = findNamedStyle(workingRequest.styleId);
      ((styleToUpdate.definition = validateNamedStyleDefinition(
        workingRequest.definition,
      )),
        (workingRequest.definition = cloneStyleValue(styleToUpdate.definition)),
        refreshBoundCells(styleToUpdate.id));
      break;
    }
    case "renameStyle": {
      const styleToRename = findNamedStyle(workingRequest.styleId);
      ((styleToRename.name = validateBoundedStyleText(
        workingRequest.name,
        "Style name",
      )),
        (workingRequest.name = styleToRename.name),
        refreshBoundCells(styleToRename.id));
      break;
    }
    case "deleteStyle":
      findNamedStyle(workingRequest.styleId);
      const styleId = workingRequest.styleId;
      ((workingCatalog.styles = workingCatalog.styles.filter(
        (style) => style.id !== styleId,
      )),
        refreshBoundCells(workingRequest.styleId),
        warnings.push(
          "Cells using this style will show default formatting plus their direct overrides. Detach the style first to retain its current appearance.",
        ));
      break;
    case "applyStyle":
    case "reapplyStyle":
    case "detachStyle": {
      const sheetOrThrow = findSheetOrThrow(workbook, workingRequest.sheetId);
      (assertSheetUnprotected(sheetOrThrow),
        workingRequest.kind !== "detachStyle" &&
          findNamedStyle(workingRequest.styleId));
      for (const selectedAddress of enumerateSelectionCells(
        workingRequest.ranges,
      )) {
        const existingCell = sheetOrThrow.cells[selectedAddress] ?? {
          raw: null,
        };
        if (workingRequest.kind === "detachStyle") {
          if (!existingCell.style?.binding) continue;
          const effectiveCellStyle = resolveEffectiveCellStyle(
              existingCell.style,
            ),
            detachedCell = { ...existingCell, style: effectiveCellStyle };
          (Object.keys(effectiveCellStyle).length || delete detachedCell.style,
            recordCellStyleChange(
              sheetOrThrow.id,
              selectedAddress,
              detachedCell,
            ));
        } else {
          const namedStyleBinding = resolveNamedStyleBinding(workingCatalog, {
            namespace: workingCatalog.namespace,
            styleId: workingRequest.styleId,
          });
          recordCellStyleChange(sheetOrThrow.id, selectedAddress, {
            ...existingCell,
            style: {
              ...(workingRequest.kind === "applyStyle"
                ? existingCell.style
                : {}),
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
  const validatedCatalog = validateWorkbookStyleCatalog(workingCatalog),
    catalogAfter =
      workingRequest.kind === "detachStyle" && !styleCatalog
        ? undefined
        : validatedCatalog,
    summaryByKind = {
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
    request: workingRequest,
    ...(styleCatalog ? { catalogBefore: cloneStyleValue(styleCatalog) } : {}),
    ...(catalogAfter ? { catalogAfter: catalogAfter } : {}),
    changes: changes,
    count: changes.length,
    summary: summaryByKind[workingRequest.kind],
    warnings: warnings,
  };
}
function assertSheetWritable(workbook, sheetId, options) {
  (options.readOnly && throwWithMessage("This workbook is read-only."),
    assertSheetUnprotected(findSheetOrThrow(workbook, sheetId)));
}
function nQ(
  workbook,
  sheetId,
  selectionAnchorAddress,
  commentText,
  authorName = "Reviewer",
  options = {},
) {
  assertSheetWritable(workbook, sheetId, options);
  const sheetOrThrow = findSheetOrThrow(workbook, sheetId),
    reviewTimestamp = resolveReviewTimestamp(options),
    reviewAuthor = normalizeReviewAuthor(authorName),
    reviewThread = {
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
            commentText,
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
    reviewThread,
  ]);
}
function findReviewThread(workbook, reviewThreadId, options) {
  options.readOnly && throwWithMessage("This workbook is read-only.");
  const matchingThread = workbook.reviewThreads?.find(
    (reviewThread) => reviewThread.id === reviewThreadId,
  );
  return matchingThread
    ? (matchingThread.orphaned &&
        throwWithMessage(
          "This discussion is retained for review because its target was deleted.",
        ),
      assertSheetWritable(workbook, matchingThread.anchor.sheetId, options),
      matchingThread)
    : throwWithMessage(
        "The discussion could not be found. Reopen the review tools.",
      );
}
function addReviewThreadReply(
  workbook,
  reviewThreadId,
  replyText,
  authorName = "Reviewer",
  options = {},
) {
  const reviewThread = findReviewThread(workbook, reviewThreadId, options);
  reviewThread.resolved &&
    throwWithMessage("Reopen the discussion before adding a reply.");
  const reviewTimestamp = resolveReviewTimestamp(
      options,
      reviewThread.updatedAt,
    ),
    updatedThread = {
      ...reviewThread,
      updatedAt: reviewTimestamp,
      messages: [
        ...reviewThread.messages,
        {
          id: crypto.randomUUID(),
          author: normalizeReviewAuthor(authorName),
          text: normalizeBoundedText(
            replyText,
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
    workbook.reviewThreads.map((reviewThread) =>
      reviewThread.id === reviewThreadId ? updatedThread : reviewThread,
    ),
  );
}
function setReviewThreadResolved(
  workbook,
  reviewThreadId,
  shouldResolve,
  authorName = "Reviewer",
  options = {},
) {
  typeof shouldResolve != "boolean" &&
    throwWithMessage("Choose whether to resolve or reopen the discussion.");
  const reviewThread = findReviewThread(workbook, reviewThreadId, options);
  if (!!reviewThread.resolved === shouldResolve) return workbook;
  const reviewTimestamp = resolveReviewTimestamp(
      options,
      reviewThread.updatedAt,
    ),
    updatedThread = { ...reviewThread, updatedAt: reviewTimestamp };
  return (
    shouldResolve
      ? (updatedThread.resolved = {
          author: normalizeReviewAuthor(authorName),
          createdAt: reviewTimestamp,
        })
      : (normalizeReviewAuthor(authorName), delete updatedThread.resolved),
    attachReviewThreadsToWorkbook(
      workbook,
      workbook.reviewThreads.map((reviewThread) =>
        reviewThread.id === reviewThreadId ? updatedThread : reviewThread,
      ),
    )
  );
}
function setCellHyperlink(
  workbook,
  sheetId,
  selectionAnchorAddress,
  hyperlink,
  options = {},
) {
  assertSheetWritable(workbook, sheetId, options);
  const sheetOrThrow = findSheetOrThrow(workbook, sheetId),
    cellAddress = normalizeCellAddress(selectionAnchorAddress),
    existingCell = sheetOrThrow.cells[cellAddress],
    cellPosition = _t(cellAddress);
  sheetOrThrow.merges?.some((merge) => {
    const rangeAddress = parseRangeAddress(merge);
    return (
      cellPosition.row >= rangeAddress.firstRow &&
      cellPosition.row <= rangeAddress.lastRow &&
      cellPosition.col >= rangeAddress.firstCol &&
      cellPosition.col <= rangeAddress.lastCol &&
      (cellPosition.row !== rangeAddress.firstRow ||
        cellPosition.col !== rangeAddress.firstCol)
    );
  }) &&
    throwWithMessage(
      "Choose the merged area's first cell or unmerge it before changing its link.",
    );
  const validatedHyperlink =
    hyperlink === null ? undefined : validateHyperlinkRecord(hyperlink);
  if (
    (validatedHyperlink?.kind === "internal" &&
      findSheetOrThrow(workbook, validatedHyperlink.sheetId),
    JSON.stringify(existingCell?.hyperlink) ===
      JSON.stringify(validatedHyperlink))
  )
    return workbook;
  !existingCell &&
    workbook.sheets.reduce(
      (total, workbookSheet) => total + Object.keys(workbookSheet.cells).length,
      0,
    ) >= workbookLimits.maxCells &&
    throwWithMessage(
      "Adding this link would exceed the workbook populated-cell limit.",
    );
  const cell = { ...(existingCell ?? { raw: null }) };
  validatedHyperlink
    ? (cell.hyperlink = validatedHyperlink)
    : delete cell.hyperlink;
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
function collectHyperlinkEntries(workbook, previousSheetId, sheetId) {
  const sheetOrThrow = findSheetOrThrow(workbook, sheetId);
  let hasRetargetedLinks = false;
  const cells = Object.fromEntries(
    Object.entries(sheetOrThrow.cells).map(([cellAddress, cellRecord]) =>
      cellRecord.hyperlink?.kind !== "internal" ||
      cellRecord.hyperlink.sheetId !== previousSheetId
        ? [cellAddress, cellRecord]
        : ((hasRetargetedLinks = true),
          [
            cellAddress,
            {
              ...cellRecord,
              hyperlink: { ...cellRecord.hyperlink, sheetId: sheetId },
            },
          ]),
    ),
  );
  return hasRetargetedLinks
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
    selectedAddresses = validateSelectionBounds(
      sheetOrThrow,
      activeCellAddress,
    );
  if (selectedAddresses.length !== 1)
    throw new Error("Select one empty total cell.");
  activeCellAddress = selectedAddresses[0];
  const cellPosition = _t(activeCellAddress),
    cell = sheetOrThrow.cells[activeCellAddress];
  if (cell?.raw !== undefined && cell.raw !== null && cell.raw !== "")
    throw new Error("Select an empty cell for AutoSum.");
  const isNumericOrFormulaCell = (rowIndex, columnIndex) => {
    if (rowIndex < 0 || columnIndex < 0) return false;
    const cell =
      sheetOrThrow.cells[`${columnIndexToLetters(columnIndex)}${rowIndex + 1}`];
    return (
      typeof cell?.raw == "number" ||
      (typeof cell?.raw == "string" &&
        cell.raw.startsWith("=") &&
        !cell.literal)
    );
  };
  let rangeStart, rangeEnd;
  if (isNumericOrFormulaCell(cellPosition.row - 1, cellPosition.col)) {
    let topRowIndex = cellPosition.row - 1;
    for (; isNumericOrFormulaCell(topRowIndex - 1, cellPosition.col);)
      topRowIndex--;
    ((rangeStart = `${columnIndexToLetters(cellPosition.col)}${topRowIndex + 1}`),
      (rangeEnd = `${columnIndexToLetters(cellPosition.col)}${cellPosition.row}`));
  } else if (isNumericOrFormulaCell(cellPosition.row, cellPosition.col - 1)) {
    let leftColumnIndex = cellPosition.col - 1;
    for (; isNumericOrFormulaCell(cellPosition.row, leftColumnIndex - 1);)
      leftColumnIndex--;
    ((rangeStart = `${columnIndexToLetters(leftColumnIndex)}${cellPosition.row + 1}`),
      (rangeEnd = `${columnIndexToLetters(cellPosition.col - 1)}${cellPosition.row + 1}`));
  } else
    throw new Error(
      "AutoSum needs adjacent numbers or formulas above or to the left.",
    );
  const totalCell = cloneCellRecord(cell);
  return (
    (totalCell.raw = `=SUM(${rangeStart}:${rangeEnd})`),
    delete totalCell.literal,
    { sheetId: sheetId, address: activeCellAddress, cell: totalCell }
  );
}
function setTableColumnDataType(workbook, sheetId, tableId, columnId) {
  const sheetOrThrow = findSheetOrThrow(workbook, sheetId),
    table = sheetOrThrow.tables?.find((table) => table.id === tableId);
  if (!table) throw Error("Table not found.");
  const columnIndex = table.columns.findIndex(
    (column) => column.id === columnId,
  );
  if (columnIndex < 0) throw Error("Select an existing table column.");
  const rangeAddress = parseRangeAddress(table.range),
    total = Math.max(
      0,
      rangeAddress.lastRow - rangeAddress.firstRow - Number(table.showTotals),
    ),
    workbookWithSettings = calculateWorkbookWithSettings(workbook),
    exceptions = [];
  for (
    let dataRowIndex = 1;
    dataRowIndex <= Math.min(total, 500);
    dataRowIndex++
  ) {
    const address = `${columnIndexToLetters(rangeAddress.firstCol + columnIndex)}${rangeAddress.firstRow + dataRowIndex + 1}`,
      cellValue = workbookWithSettings(sheetId, address);
    coerceImportedValue(cellValue, table.columns[columnIndex].dataType) ||
      exceptions.push({ address: address, value: cellValue });
  }
  return {
    dataType: table.columns[columnIndex].dataType ?? "any",
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
  const matchingTable = sheetOrThrow.tables?.find((table) =>
    Dp(
      parseRangeAddress(table.range),
      rangeAddress.firstRow,
      rangeAddress.firstCol,
    ),
  );
  return matchingTable
    ? matchingTable.range
    : formatOccupiedRangeAddress(sheetOrThrow, false);
}
function suggestTableHeaderName(workbook, sheetId, range) {
  const sheetOrThrow = findSheetOrThrow(workbook, sheetId),
    sourceRange = resolveSourceRange(range, sheetOrThrow),
    usedNames = new Set(),
    suggestions = [];
  for (
    let firstCol = sourceRange.firstCol;
    firstCol <= sourceRange.lastCol;
    firstCol++
  ) {
    const address = Cn(sourceRange.firstRow, firstCol),
      cell = sheetOrThrow.cells[address],
      original = cell?.raw ?? null;
    let baseName = `Column${firstCol - sourceRange.firstCol + 1}`;
    try {
      isNullOrEmptyValue(original) ||
        (baseName = validateColumnName(String(original)));
    } catch {}
    let suggested = baseName,
      duplicateSuffix = 2;
    for (; usedNames.has(suggested.toLowerCase());)
      suggested = `${baseName.slice(0, 250)}${duplicateSuffix++}`;
    (usedNames.add(suggested.toLowerCase()),
      suggestions.push({
        address: address,
        original: original,
        suggested: suggested,
        requiresResolution:
          original !== suggested ||
          (suggested.startsWith("=") && !cell?.literal),
      }));
  }
  return suggestions;
}
function validateReviewedTableHeaders(workbook, sheetId, range, headers) {
  const suggestions = suggestTableHeaderName(workbook, sheetId, range);
  if (headers === undefined)
    return (
      suggestions.some((suggestion) => suggestion.requiresResolution) &&
        throwWorkbookError(
          "Review table headers: resolve blank, duplicate, non-text or formula headers explicitly. Original header values will be retained in native source evidence.",
        ),
      suggestions.map((headerSuggestion) => headerSuggestion.suggested)
    );
  (!Array.isArray(headers) || headers.length !== suggestions.length) &&
    throwWorkbookError("Provide one reviewed header for every table column.");
  const validatedHeaders = headers.map(validateColumnName);
  return (
    new Set(validatedHeaders.map((header) => header.toLowerCase())).size !==
      validatedHeaders.length &&
      throwWorkbookError(
        "Reviewed table headers must be unique, nonblank text.",
      ),
    validatedHeaders
  );
}
function renameTableHeaderField(workbook, nextWorkbook) {
  const renames = [];
  for (const sheet of workbook.sheets) {
    const nextSheet = nextWorkbook.sheets.find(
      (candidateSheet) => candidateSheet.id === sheet.id,
    );
    if (nextSheet)
      for (const old of sheet.tables ?? []) {
        const nextTable = nextSheet.tables?.find(
          (table) => table.id === old.id,
        );
        if (!nextTable) continue;
        const rangeAddress = parseRangeAddress(old.range),
          nextRangeAddress = parseRangeAddress(nextTable.range),
          names = {},
          columns = nextTable.columns.map((column, nextColumnIndex) => {
            const oldColumnIndex = old.columns.findIndex(
              (oldColumnEntry) => oldColumnEntry.id === column.id,
            );
            if (oldColumnIndex < 0) return column;
            const oldColumn = old.columns[oldColumnIndex],
              cell =
                sheet.cells[
                  Cn(
                    rangeAddress.firstRow,
                    rangeAddress.firstCol + oldColumnIndex,
                  )
                ],
              nextCell =
                nextSheet.cells[
                  Cn(
                    nextRangeAddress.firstRow,
                    nextRangeAddress.firstCol + nextColumnIndex,
                  )
                ],
              nextName =
                cell?.raw !== nextCell?.raw ||
                cell?.literal !== nextCell?.literal
                  ? nextCell && !isNullOrEmptyValue(nextCell.raw)
                    ? validateColumnName(String(nextCell.raw))
                    : throwWorkbookError(
                        "Table headers must be nonblank text. The header edit was not applied.",
                      )
                  : column.name;
            return (
              nextName !== oldColumn.name &&
                Object.defineProperty(names, oldColumn.name, {
                  value: nextName,
                  enumerable: true,
                }),
              nextName === column.name ? column : { ...column, name: nextName }
            );
          });
        (old.name !== nextTable.name || Object.keys(names).length) &&
          (new Set(columns.map((column) => column.name.toLowerCase())).size !==
            columns.length &&
            throwWorkbookError(
              "Table headers must be unique, nonblank text. The header edit was not applied.",
            ),
          renames.push({
            sheetId: nextSheet.id,
            old: old,
            current: nextTable,
            columns: columns,
            names: names,
          }));
      }
  }
  if (!renames.length) return nextWorkbook;
  const updatedWorkbook = structuredClone(nextWorkbook);
  for (const rename of renames) {
    const sheetOrThrow = findSheetOrThrow(updatedWorkbook, rename.sheetId),
      table = sheetOrThrow.tables.find(
        (table) => table.id === rename.current.id,
      ),
      renameField = (fieldName) =>
        Object.hasOwn(rename.names, fieldName)
          ? rename.names[fieldName]
          : fieldName;
    ((table.columns = rename.columns),
      (table.totals = Object.fromEntries(
        Object.entries(table.totals).map(([totalFieldName, totalValue]) => [
          renameField(totalFieldName),
          totalValue,
        ]),
      )));
    const rangeAddress = parseRangeAddress(table.range);
    table.columns.forEach((column, columnIndex) => {
      const headerCellAddress = Cn(
        rangeAddress.firstRow,
        rangeAddress.firstCol + columnIndex,
      );
      sheetOrThrow.cells[headerCellAddress] = {
        ...(sheetOrThrow.cells[headerCellAddress] ?? { raw: column.name }),
        raw: column.name,
        ...(column.name.startsWith("=")
          ? { literal: true }
          : { literal: undefined }),
      };
    });
    const refersToTable = (sourceReference) =>
      sourceReference.tableId === table.id ||
      (!sourceReference.tableId &&
        sourceReference.sheetId === sheetOrThrow.id &&
        toA1RangeAddress(sourceReference.range) === rename.old.range);
    for (const sheet of updatedWorkbook.sheets) {
      for (const sheetTable of sheet.tables ?? [])
        for (const column of sheetTable.columns)
          column.calculatedFormula &&= $v(
            column.calculatedFormula,
            rename.old.name,
            table.name,
            rename.names,
            sheetTable.id === table.id && sheet.id === sheetOrThrow.id,
          );
      for (const [cellAddress, cell] of Object.entries(sheet.cells)) {
        if (
          typeof cell.raw != "string" ||
          !cell.raw.startsWith("=") ||
          cell.literal
        )
          continue;
        const cellPosition = _t(cellAddress),
          isInsideTable =
            sheet.id === sheetOrThrow.id &&
            Dp(rangeAddress, cellPosition.row, cellPosition.col);
        cell.raw = $v(
          cell.raw,
          rename.old.name,
          table.name,
          rename.names,
          isInsideTable,
        );
      }
      for (const conditionalFormat of sheet.conditionalFormats ?? [])
        if (conditionalFormat.kind === "formula" && conditionalFormat.formula) {
          const formatRangeAddress = parseRangeAddress(conditionalFormat.range),
            formulaOutsideTable = $v(
              conditionalFormat.formula,
              rename.old.name,
              table.name,
              rename.names,
              false,
            ),
            formulaInsideTable = $v(
              conditionalFormat.formula,
              rename.old.name,
              table.name,
              rename.names,
              true,
            ),
            overlapsTable =
              sheet.id === sheetOrThrow.id &&
              rangesIntersect(rangeAddress, formatRangeAddress),
            isFullyInsideTable =
              overlapsTable &&
              Dp(
                rangeAddress,
                formatRangeAddress.firstRow,
                formatRangeAddress.firstCol,
              ) &&
              Dp(
                rangeAddress,
                formatRangeAddress.lastRow,
                formatRangeAddress.lastCol,
              );
          (formulaOutsideTable !== formulaInsideTable &&
            overlapsTable &&
            !isFullyInsideTable &&
            throwWorkbookError(
              "This conditional formula uses unqualified table fields across a table boundary. Split or remove its applied range before renaming the field.",
            ),
            (conditionalFormat.formula = isFullyInsideTable
              ? formulaInsideTable
              : formulaOutsideTable));
        }
      for (const pivotOrQuery of [
        ...(sheet.pivots ?? []),
        ...(sheet.queries ?? []),
      ])
        if (refersToTable(pivotOrQuery.source))
          if (
            ((pivotOrQuery.source = {
              ...pivotOrQuery.source,
              tableId: table.id,
              range: table.range,
            }),
            (pivotOrQuery.filters &&= pivotOrQuery.filters.map((filter) => ({
              ...filter,
              field: renameField(filter.field),
            }))),
            "rowFields" in pivotOrQuery)
          ) {
            ((pivotOrQuery.rowFields = pivotOrQuery.rowFields.map(renameField)),
              (pivotOrQuery.columnFields =
                pivotOrQuery.columnFields.map(renameField)),
              (pivotOrQuery.values = pivotOrQuery.values.map((value) => ({
                ...value,
                field: renameField(value.field),
              }))),
              (pivotOrQuery.dateGroups &&= pivotOrQuery.dateGroups.map(
                (dateGroup) => ({
                  ...dateGroup,
                  field: renameField(dateGroup.field),
                }),
              )));
            for (const sectionKey of ["slicers", "timelines", "sorts"])
              pivotOrQuery[sectionKey] &&= pivotOrQuery[sectionKey].map(
                (sectionItem) => ({
                  ...sectionItem,
                  field: renameField(sectionItem.field),
                }),
              );
            pivotOrQuery.calculatedFields &&= pivotOrQuery.calculatedFields.map(
              (calculatedField) => ({
                ...calculatedField,
                field: renameField(calculatedField.field),
                ...(calculatedField.otherField
                  ? { otherField: renameField(calculatedField.otherField) }
                  : {}),
              }),
            );
          } else
            ((pivotOrQuery.fields = pivotOrQuery.fields.map(renameField)),
              (pivotOrQuery.rename &&= Object.fromEntries(
                Object.entries(pivotOrQuery.rename).map(
                  ([renamedField, newName]) => [
                    renameField(renamedField),
                    newName,
                  ],
                ),
              )),
              (pivotOrQuery.sort &&= {
                ...pivotOrQuery.sort,
                field: renameField(pivotOrQuery.sort.field),
              }),
              (pivotOrQuery.derived &&= pivotOrQuery.derived.map(
                (derivedItem) => ({
                  ...derivedItem,
                  field: renameField(derivedItem.field),
                  ...(derivedItem.otherField
                    ? { otherField: renameField(derivedItem.otherField) }
                    : {}),
                }),
              )));
    }
    for (const definedName of updatedWorkbook.definedNames ?? [])
      definedName.refersTo = $v(
        definedName.refersTo,
        rename.old.name,
        table.name,
        rename.names,
      );
    normalizeTableDefinitions(sheetOrThrow, sheetOrThrow);
  }
  assertUniqueTableNames(updatedWorkbook);
  for (const sheet of workbook.sheets) {
    if (!sheet.protection?.enabled) continue;
    const updatedSheet = updatedWorkbook.sheets.find(
      (sheetCandidate) => sheetCandidate.id === sheet.id,
    );
    JSON.stringify(sheet) !== JSON.stringify(updatedSheet) &&
      throwWorkbookError(
        `This change would update dependent cells or analysis on protected worksheet ${sheet.name}. Unprotect it before renaming the table field.`,
      );
  }
  return updatedWorkbook;
}
function buildPivotSource(workbook, source) {
  const sourceReference = normalizeSourceReference(source),
    sheetOrThrow = findSheetOrThrow(workbook, sourceReference.sheetId),
    sourceTable = sourceReference.tableId
      ? (sheetOrThrow.tables?.find(
          (table) => table.id === sourceReference.tableId,
        ) ?? throwWorkbookError("The source table no longer exists."))
      : sheetOrThrow.tables?.find(
          (table) => table.range === toA1RangeAddress(sourceReference.range),
        ),
    sourceRange = resolveSourceRange(
      sourceTable?.range ?? sourceReference.range,
      sheetOrThrow,
    ),
    workbookWithSettings = calculateWorkbookWithSettings(workbook);
  sourceTable?.showTotals && sourceRange.lastRow--;
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
  for (
    let rowIndex = sourceRange.firstRow + 1;
    rowIndex <= sourceRange.lastRow;
    rowIndex++
  ) {
    const values = fields.map((field, fieldIndex) =>
      workbookWithSettings(
        sheetOrThrow.id,
        Cn(rowIndex, sourceRange.firstCol + fieldIndex),
      ),
    );
    (values.forEach((value, columnOffset) => {
      const cellAddress = Cn(rowIndex, sourceRange.firstCol + columnOffset);
      typeof value == "string" &&
        value.startsWith("#") &&
        workbookWithSettings.evaluateFormula(
          sheetOrThrow.id,
          cellAddress,
          `=ISERROR(${cellAddress})`,
        ) === true &&
        throwWorkbookError(
          `Source calculation error ${value} at ${sheetOrThrow.name}!${cellAddress}. Repair the source calculation before building or refreshing analysis; no totals were written.`,
        );
    }),
      (values.some((value) => !isNullOrEmptyValue(value)) ||
        fields.some(
          (field, sourceColumnOffset) =>
            !isNullOrEmptyValue(
              sheetOrThrow.cells[
                Cn(rowIndex, sourceRange.firstCol + sourceColumnOffset)
              ]?.raw ?? null,
            ),
        )) &&
        rows.push({ index: rowIndex, values: values }));
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
        range: sourceTable?.range ?? toA1RangeAddress(sourceReference.range),
        ...(sourceTable ? { tableId: sourceTable.id } : {}),
      },
    }
  );
}
function getPivotSourceFields(workbook, sheetId, range) {
  return buildPivotSource(workbook, { sheetId: sheetId, range: range }).fields;
}
function aggregatePivotRows(workbook, definition, pivotSource) {
  const sheetOrThrow = findSheetOrThrow(workbook, pivotSource.source.sheetId),
    rangeAddress = parseRangeAddress(pivotSource.source.range),
    cells = Object.entries(sheetOrThrow.cells)
      .filter(([cellAddress]) => {
        const position = _t(cellAddress);
        return Dp(rangeAddress, position.row, position.col);
      })
      .sort(([leftAddress], [rightAddress]) =>
        leftAddress.localeCompare(rightAddress),
      )
      .map(([address, cell]) => [address, cell.raw, !!cell.literal]),
    { destination, ...settings } = normalizePivotDefinition(definition);
  return computePivotRefreshFingerprint({
    source: pivotSource.source,
    fields: pivotSource.fields,
    rows: pivotSource.rows,
    cells: cells,
    settings: settings,
  });
}
function getPivotRefreshState(workbook, pivotId) {
  const sheet = workbook.sheets.find((sheet) =>
      sheet.pivots?.some((pivot) => pivot.id === pivotId),
    ),
    definition = sheet?.pivots?.find((pivot) => pivot.id === pivotId);
  if (
    ((!sheet || !definition) &&
      throwWorkbookError("Pivot definition not found."),
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
    collectRangeValues(sheet, definition.outputRange) !==
    definition.refresh.outputFingerprint;
  try {
    const pivotSource = buildPivotSource(workbook, definition.source);
    validatePivotFieldReferences(definition, pivotSource.fields);
    const sourceChanged =
      aggregatePivotRows(workbook, definition, pivotSource) !==
      definition.refresh.sourceFingerprint;
    return {
      state: sourceChanged || outputModified ? "stale" : "current",
      reason: outputModified
        ? "Output cells were edited. Refresh into a new result before relying on these totals."
        : sourceChanged
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
function inspectPivotControlFields(workbook, pivotId, fieldNames) {
  fieldNames.length > 30 &&
    throwWorkbookError("Inspect at most 30 pivot control fields together.");
  const { definition } = findPivotDefinition(workbook, pivotId),
    pivotSource = buildPivotSource(workbook, definition.source);
  return new Map(
    [...new Set(fieldNames)].map((fieldName) => {
      const fieldIndex = pivotSource.fields.indexOf(fieldName),
        itemsByKey = new Map();
      if (fieldIndex < 0)
        return [
          fieldName,
          {
            items: [],
            error: `Broken pivot field “${fieldName}”: source field was not found. Restore or reselect it.`,
          },
        ];
      for (const row of pivotSource.rows)
        if (
          (itemsByKey.set(
            buildSlicerItemKey(row.values[fieldIndex]),
            row.values[fieldIndex],
          ),
          itemsByKey.size > 500)
        )
          return [
            fieldName,
            {
              items: [],
              error:
                "This slicer exceeds 500 distinct items. Choose another field or reduce the source.",
            },
          ];
      return [fieldName, { items: [...itemsByKey.values()], error: "" }];
    }),
  );
}
function getPivotSlicerItems(workbook, pivotId, field) {
  const fieldResult = inspectPivotControlFields(workbook, pivotId, [field]).get(
    field,
  );
  return (
    fieldResult.error && throwWorkbookError(fieldResult.error),
    fieldResult.items
  );
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
    valueIndex = definition.chart?.valueIndex ?? 0;
  built.chartRows.length > 200 &&
    throwWorkbookError(
      "PivotCharts are limited to 200 leaf groups. Filter or regroup before charting.",
    );
  const points = built.chartRows.map((chartRow) => ({
    label: chartRow.label,
    value: chartRow.values[valueIndex],
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
      measure: built.measures[valueIndex].name,
      showAs: definition.values[valueIndex].showAs ?? "value",
      points: points,
      sourceRevision: definition.sourceRevision,
    }
  );
}
function buildPivotDrillDown(workbook, pivotId, cellAddress) {
  const { definition, config, data, rows, built } = buildPivotModel(
      workbook,
      pivotId,
    ),
    cellPosition = _t(cellAddress),
    rangeAddress = parseRangeAddress(definition.outputRange),
    rowAxisIndex = cellPosition.row - rangeAddress.firstRow - 1,
    valueColumnOffset =
      cellPosition.col -
      rangeAddress.firstCol -
      Math.max(1, config.rowFields.length);
  (rowAxisIndex < 0 ||
    rowAxisIndex >= built.rowAxes.length ||
    valueColumnOffset < 0 ||
    cellPosition.col > rangeAddress.lastCol ||
    cellPosition.row > rangeAddress.lastRow) &&
    throwWorkbookError(
      "Choose a PivotTable value cell, subtotal or grand total; headers and group labels cannot be drilled.",
    );
  const rowAxis = built.rowAxes[rowAxisIndex],
    columnAxis =
      built.columnAxes[Math.floor(valueColumnOffset / config.values.length)],
    valueIndex = valueColumnOffset % config.values.length,
    matchesAxis = (pivotRow, axisGroup, groupFields) =>
      !axisGroup ||
      axisGroup.parts.every((part, partIndex) => {
        const fieldValue =
            pivotRow.values[
              findSourceFieldIndex(data.fields, groupFields[partIndex])
            ],
          matchingDateGroup = config.dateGroups?.find(
            (dateGroup) => dateGroup.field === groupFields[partIndex],
          );
        return (
          JSON.stringify(
            matchingDateGroup
              ? bucketPivotDateValue(
                  fieldValue,
                  matchingDateGroup,
                  pivotRow.index,
                )
              : isNullOrEmptyValue(fieldValue)
                ? null
                : fieldValue,
          ) === JSON.stringify(part)
        );
      }),
    matchingRows = rows.filter(
      (row) =>
        matchesAxis(row, rowAxis, config.rowFields) &&
        matchesAxis(row, columnAxis, config.columnFields),
    ),
    sheetOrThrow = findSheetOrThrow(workbook, data.source.sheetId),
    sourceRangeAddress = parseRangeAddress(data.source.range),
    records = matchingRows.map((matchingRow) => ({
      locator: {
        workbookId: workbook.id,
        sheetId: sheetOrThrow.id,
        sheetName: sheetOrThrow.name,
        row: matchingRow.index + 1,
        range: `${columnIndexToLetters(sourceRangeAddress.firstCol)}${matchingRow.index + 1}:${columnIndexToLetters(sourceRangeAddress.lastCol)}${matchingRow.index + 1}`,
        original: findRowOrigin(workbook, sheetOrThrow.id, matchingRow.index),
      },
      values: [...matchingRow.values],
      cells: data.fields.map((field, fieldIndex) => ({
        address: Cn(
          matchingRow.index,
          sourceRangeAddress.firstCol + fieldIndex,
        ),
        cell: structuredClone(
          sheetOrThrow.cells[
            Cn(matchingRow.index, sourceRangeAddress.firstCol + fieldIndex)
          ] ?? { raw: null },
        ),
      })),
    }));
  return {
    context: buildPivotStatusRecord(workbook, definition, {
      state: "current",
      reason: "Source and output fingerprints match the saved calculation.",
    }),
    address: `${columnIndexToLetters(cellPosition.col)}${cellPosition.row + 1}`,
    valueIndex: valueIndex,
    measure: config.values[valueIndex],
    underlyingValue:
      built.baseMatrix[rowAxisIndex + 1][
        cellPosition.col - rangeAddress.firstCol
      ],
    displayedValue:
      built.matrix[rowAxisIndex + 1][cellPosition.col - rangeAddress.firstCol],
    rowGroup: rowAxis?.parts ?? null,
    columnGroup: columnAxis?.parts ?? null,
    fields: [...data.fields],
    records: records,
  };
}
function validateTotalsRowForAppend(workbook, sheetId, table) {
  const sheetOrThrow = findSheetOrThrow(workbook, sheetId),
    rangeAddress = parseRangeAddress(table.range),
    assertFormulaSafe = (formulaText, formulaSheetId) => {
      const formulaWithoutStrings = stripFormulaStringLiterals(
        formulaText.replace(/"(?:[^"]|"")*"/g, '""'),
      );
      /\b(?:INDIRECT|OFFSET|EVALUATE|LET|LAMBDA)\s*\(/i.test(
        formulaWithoutStrings,
      ) &&
        throwWorkbookError(
          "A dynamic formula could refer to the moving totals row. Preserve or remove that ambiguity before appending.",
        );
      for (const rowRangeMatch of formulaWithoutStrings.matchAll(
        new RegExp(
          "(?<![A-Za-z0-9_.])(?:(?:'((?:[^']|'')+)'|([A-Za-z_\\\\][A-Za-z0-9_.\\\\]*))\\s*!\\s*)?(\\$?[1-9]\\d*)\\s*:\\s*(\\$?[1-9]\\d*)(?![A-Za-z0-9_.])",
          "g",
        ),
      )) {
        const matchedSheetName =
          rowRangeMatch[1]?.replace(/''/g, "'") ?? rowRangeMatch[2];
        if (
          matchedSheetName
            ? matchedSheetName.toLowerCase() !== sheetOrThrow.name.toLowerCase()
            : formulaSheetId !== undefined && formulaSheetId !== sheetId
        )
          continue;
        const startRow = Number(rowRangeMatch[3].replace("$", "")) - 1,
          endRow = Number(rowRangeMatch[4].replace("$", "")) - 1;
        rangeAddress.lastRow >= Math.min(startRow, endRow) &&
          rangeAddress.lastRow <= Math.max(startRow, endRow) &&
          throwWorkbookError(
            "A formula refers directly to the moving totals row. Use a structured table totals reference or move those references before appending.",
          );
      }
      for (const cellRefMatch of formulaWithoutStrings.matchAll(
        new RegExp(
          "(?<![A-Za-z0-9_.])(?:(?:'((?:[^']|'')+)'|([A-Za-z_\\\\][A-Za-z0-9_.\\\\]*))\\s*!\\s*)?(\\$?[A-Za-z]{1,3}\\$?[1-9]\\d*)(?:\\s*:\\s*(\\$?[A-Za-z]{1,3}\\$?[1-9]\\d*))?(?![A-Za-z0-9_.])",
          "g",
        ),
      )) {
        const cellSheetName =
          cellRefMatch[1]?.replace(/''/g, "'") ?? cellRefMatch[2];
        if (
          cellSheetName
            ? cellSheetName.toLowerCase() !== sheetOrThrow.name.toLowerCase()
            : formulaSheetId !== undefined && formulaSheetId !== sheetId
        )
          continue;
        const cellRefRange = parseRangeAddress(
          `${cellRefMatch[3]}:${cellRefMatch[4] ?? cellRefMatch[3]}`,
        );
        rangesIntersect(cellRefRange, {
          ...rangeAddress,
          firstRow: rangeAddress.lastRow,
        }) &&
          throwWorkbookError(
            "A formula refers directly to the moving totals row. Use a structured table totals reference or move those references before appending.",
          );
      }
    };
  for (const sheet of workbook.sheets) {
    for (const [cellAddress, cell] of Object.entries(sheet.cells)) {
      const cellPosition = _t(cellAddress);
      if (
        !(
          sheet.id === sheetId &&
          cellPosition.row === rangeAddress.lastRow &&
          cellPosition.col >= rangeAddress.firstCol &&
          cellPosition.col <= rangeAddress.lastCol
        ) &&
        (typeof cell.raw == "string" &&
          cell.raw.startsWith("=") &&
          !cell.literal &&
          assertFormulaSafe(cell.raw, sheet.id),
        cell.hyperlink?.kind === "internal" &&
          cell.hyperlink.sheetId === sheetId)
      ) {
        const hyperlinkTarget = _t(cell.hyperlink.address);
        hyperlinkTarget.row === rangeAddress.lastRow &&
          hyperlinkTarget.col >= rangeAddress.firstCol &&
          hyperlinkTarget.col <= rangeAddress.lastCol &&
          throwWorkbookError(
            "An internal hyperlink targets the moving totals row. Retarget or remove it before appending.",
          );
      }
    }
    for (const conditionalFormat of sheet.conditionalFormats ?? [])
      conditionalFormat.kind === "formula" &&
        conditionalFormat.formula &&
        assertFormulaSafe(conditionalFormat.formula, sheet.id);
    for (const sheetTable of sheet.tables ?? [])
      for (const column of sheetTable.columns)
        column.calculatedFormula &&
          assertFormulaSafe(column.calculatedFormula, sheet.id);
  }
  for (const definedName of workbook.definedNames ?? [])
    assertFormulaSafe(definedName.refersTo, definedName.scopeSheetId);
  for (
    let firstCol = rangeAddress.firstCol;
    firstCol <= rangeAddress.lastCol;
    firstCol++
  ) {
    const cell = sheetOrThrow.cells[Cn(rangeAddress.lastRow, firstCol)],
      tableTotalCell = buildTableTotalCell(
        table,
        firstCol - rangeAddress.firstCol,
      );
    cell &&
      (cell.raw !== tableTotalCell.raw || cell.literal) &&
      throwWorkbookError(
        "The totals row was edited. Preserve its values before appending.",
      );
  }
}
function resizeTableRange(workbook, sheetId, table, sourceRange, showTotals) {
  const sheetOrThrow = findSheetOrThrow(workbook, sheetId),
    rangeAddress = parseRangeAddress(table.range);
  (showTotals !== table.showTotals &&
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
      (table.showTotals ? table.columns.length : 0) >
      tableMutationLimits.changedCells &&
      throwWorkbookError("Table append/resize may change at most 500 cells."));
  for (
    let rowIndex = rangeAddress.lastRow + 1;
    rowIndex <= sourceRange.lastRow;
    rowIndex++
  )
    for (
      let firstCol = sourceRange.firstCol;
      firstCol <= sourceRange.lastCol;
      firstCol++
    )
      sheetOrThrow.cells[Cn(rowIndex, firstCol)] &&
        throwWorkbookError(
          "Table append/resize would overwrite adjacent cells or their formatting. Move that content first.",
        );
  if (table.showTotals && sourceRange.lastRow !== rangeAddress.lastRow) {
    validateTotalsRowForAppend(workbook, sheetId, table);
    for (const anchoredItem of [
      ...(workbook.formulaWatches ?? []),
      ...(workbook.reviewThreads ?? []),
    ]) {
      if (anchoredItem.orphaned || anchoredItem.anchor.sheetId !== sheetId)
        continue;
      const anchorPosition = _t(anchoredItem.anchor.address);
      anchorPosition.row === sourceRange.lastRow &&
        anchorPosition.col >= sourceRange.firstCol &&
        anchorPosition.col <= sourceRange.lastCol &&
        throwWorkbookError(
          "The new totals row already has a saved watch or review discussion. Preserve or remove that destination anchor before appending.",
        );
    }
  }
}
function moveTotalsRowCells(workbook, next, sheetId, previousTable, table) {
  const rangeAddress = parseRangeAddress(previousTable.range),
    tableAddress = parseRangeAddress(table.range),
    sheetOrThrow = findSheetOrThrow(workbook, sheetId),
    nextSheet = findSheetOrThrow(next, sheetId);
  for (
    let firstCol = rangeAddress.firstCol;
    firstCol <= rangeAddress.lastCol;
    firstCol++
  ) {
    const sourceCellKey = Cn(rangeAddress.lastRow, firstCol),
      targetCellKey = Cn(tableAddress.lastRow, firstCol),
      sourceCell = sheetOrThrow.cells[sourceCellKey];
    if (sourceCell) {
      const totalCell = {
        ...structuredClone(sourceCell),
        ...buildTableTotalCell(table, firstCol - rangeAddress.firstCol),
      };
      (delete totalCell.literal, (nextSheet.cells[targetCellKey] = totalCell));
    }
  }
  const relocateAnchor = (anchor) => {
    if (anchor.sheetId !== sheetId) return;
    const anchorPosition = _t(anchor.address);
    anchorPosition.row === rangeAddress.lastRow &&
      anchorPosition.col >= rangeAddress.firstCol &&
      anchorPosition.col <= rangeAddress.lastCol &&
      (anchor.address = Cn(tableAddress.lastRow, anchorPosition.col));
  };
  for (const formulaWatch of next.formulaWatches ?? [])
    formulaWatch.orphaned || relocateAnchor(formulaWatch.anchor);
  for (const reviewThread of next.reviewThreads ?? [])
    reviewThread.orphaned || relocateAnchor(reviewThread.anchor);
}
function resolveTableDataRange(workbook, sheetId, range = "A1") {
  const sheetOrThrow = findSheetOrThrow(workbook, sheetId),
    rangeAddress = parseRangeAddress(range);
  if (
    rangeAddress.firstRow !== rangeAddress.lastRow ||
    rangeAddress.firstCol !== rangeAddress.lastCol
  )
    return range;
  const matchingTable = sheetOrThrow.tables?.find((table) => {
    const candidateTableAddress = parseRangeAddress(table.range);
    return (
      rangeAddress.firstRow >= candidateTableAddress.firstRow &&
      rangeAddress.firstRow <= candidateTableAddress.lastRow &&
      rangeAddress.firstCol >= candidateTableAddress.firstCol &&
      rangeAddress.firstCol <= candidateTableAddress.lastCol
    );
  });
  if (!matchingTable) return formatOccupiedRangeAddress(sheetOrThrow, false);
  const tableAddress = parseRangeAddress(matchingTable.range);
  return `${columnIndexToLetters(tableAddress.firstCol)}${tableAddress.firstRow + 1}:${columnIndexToLetters(tableAddress.lastCol)}${tableAddress.lastRow + 1 - (matchingTable.showTotals ? 1 : 0)}`;
}
function validateFilterRequest(workbook, sheetId, filterRequest) {
  if (
    !filterRequest ||
    typeof filterRequest.range != "string" ||
    !Array.isArray(filterRequest.columns) ||
    filterRequest.columns.length > filterEvaluationLimits.columns ||
    [filterRequest.caseSensitive, filterRequest.unique].some(
      (flagValue) => flagValue !== undefined && typeof flagValue != "boolean",
    )
  )
    throw Error("Choose a header range and up to 64 column filters.");
  const sheetOrThrow = findSheetOrThrow(workbook, sheetId),
    rangeAddress = parseRangeAddress(filterRequest.range),
    dataRowCount = rangeAddress.lastRow - rangeAddress.firstRow,
    columnCount = rangeAddress.lastCol - rangeAddress.firstCol + 1;
  if (
    rangeAddress.lastRow >= sheetOrThrow.rowCount ||
    rangeAddress.lastCol >= sheetOrThrow.colCount ||
    dataRowCount < 1
  )
    throw Error(
      "Filter range needs one header row and at least one data row within the worksheet.",
    );
  if (
    dataRowCount > filterEvaluationLimits.rows ||
    (dataRowCount + 1) * columnCount > filterEvaluationLimits.cells ||
    columnCount > filterEvaluationLimits.columns
  )
    throw Error(
      "Filtering is limited to 10,000 data rows, 100,000 cells including headers, and 64 columns. Choose a smaller range.",
    );
  if (
    sheetOrThrow.merges?.some((merge) => {
      const mergeAddress = parseRangeAddress(merge);
      return (
        mergeAddress.firstRow <= rangeAddress.lastRow &&
        mergeAddress.lastRow >= rangeAddress.firstRow &&
        mergeAddress.firstCol <= rangeAddress.lastCol &&
        mergeAddress.lastCol >= rangeAddress.firstCol
      );
    })
  )
    throw Error(
      "Unmerge cells in the filter range first; records cannot be split by filtering.",
    );
  const isFilterColumn = (columnIndex) =>
      Number.isInteger(columnIndex) &&
      columnIndex >= rangeAddress.firstCol &&
      columnIndex <= rangeAddress.lastCol,
    validateCondition = (filterCondition) => {
      if (
        !filterCondition ||
        !isFilterColumn(filterCondition.column) ||
        !["value", "text", "number", "date"].includes(filterCondition.type) ||
        !textFilterOperators.includes(filterCondition.operator) ||
        (!["blank", "notBlank"].includes(filterCondition.operator) &&
          !isStorableCellValue(filterCondition.value))
      )
        throw Error("Invalid filter condition, column or value.");
      if (
        !["blank", "notBlank"].includes(filterCondition.operator) &&
        filterCondition.type === "text" &&
        typeof filterCondition.value != "string"
      )
        throw Error("Text criteria require a text value.");
      if (
        typeof filterCondition.value == "string" &&
        filterCondition.value.length > filterEvaluationLimits.text
      )
        throw Error("Filter criteria text is limited to 1,024 characters.");
      if (
        [
          "contains",
          "notContains",
          "startsWith",
          "endsWith",
          "wildcard",
        ].includes(filterCondition.operator) &&
        filterCondition.type !== "text"
      )
        throw Error("Text matching requires a Text condition.");
      if (
        ["gt", "gte", "lt", "lte"].includes(filterCondition.operator) &&
        filterCondition.type === "value"
      )
        throw Error("Choose Number, Date or Text for ordered comparisons.");
      if (
        !["blank", "notBlank"].includes(filterCondition.operator) &&
        ["number", "date"].includes(filterCondition.type) &&
        resolveDatePolicyContext(
          filterCondition.value,
          filterCondition.type,
          workbook,
        ) === null
      )
        throw Error(
          `Invalid ${filterCondition.type} filter value. Use a finite number or a date in the workbook's date policy.`,
        );
    };
  let conditionCount = 0;
  const seenColumns = new Set();
  for (const column of filterRequest.columns) {
    if (
      !column ||
      !isFilterColumn(column.column) ||
      seenColumns.has(column.column) ||
      (column.join !== undefined && !["and", "or"].includes(column.join))
    )
      throw Error("Filter columns must be valid and distinct.");
    if (
      (seenColumns.add(column.column),
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
    for (const columnCondition of column.conditions ?? []) {
      if (columnCondition.column !== column.column)
        throw Error("A column condition must use its own column.");
      (validateCondition(columnCondition), conditionCount++);
    }
  }
  if (
    filterRequest.advanced !== undefined &&
    (!Array.isArray(filterRequest.advanced) ||
      filterRequest.advanced.length > filterEvaluationLimits.conditions)
  )
    throw Error("Invalid advanced filter groups.");
  for (const advancedGroup of filterRequest.advanced ?? []) {
    if (
      !Array.isArray(advancedGroup) ||
      !advancedGroup.length ||
      advancedGroup.length > filterEvaluationLimits.conditions
    )
      throw Error(
        "Each advanced OR group must contain at least one AND condition.",
      );
    for (const advancedCondition of advancedGroup)
      (validateCondition(advancedCondition), conditionCount++);
  }
  if (conditionCount > filterEvaluationLimits.conditions)
    throw Error("At most 128 combined filter conditions are allowed.");
  return structuredClone({
    ...filterRequest,
    range: `${columnIndexToLetters(rangeAddress.firstCol)}${rangeAddress.firstRow + 1}:${columnIndexToLetters(rangeAddress.lastCol)}${rangeAddress.lastRow + 1}`,
  });
}
function applyFilterConditions(
  workbook,
  sheetId,
  filterRequest,
  workbookWithSettings = calculateWorkbookWithSettings(workbook),
) {
  const definition = validateFilterRequest(workbook, sheetId, filterRequest),
    sheetOrThrow = findSheetOrThrow(workbook, sheetId),
    rangeAddress = parseRangeAddress(definition.range),
    isCaseSensitive = !!definition.caseSensitive,
    axisVisibilityIndex = createAxisVisibilityIndex(
      sheetOrThrow.rowCount,
      cn(sheetOrThrow, "row"),
    ),
    sourceRows = [],
    matchingSourceRows = [],
    visibleSourceRows = [],
    excludedRows = [],
    seenUniqueRows = new Set(),
    originalSourceIds = {},
    columnValueSets = definition.columns.map((column) =>
      column.values === undefined
        ? null
        : new Set(
            column.values.map((value) =>
              comparableCellValueKey(value, isCaseSensitive),
            ),
          ),
    );
  let manuallyHiddenRows = 0,
    hiddenMatchingRows = 0,
    comparisonWork = 0;
  const evaluateCondition = (cellValue, condition) => {
      if (
        ((comparisonWork +=
          condition.operator === "wildcard"
            ? Math.max(1, String(cellValue ?? "").length) *
              Math.max(1, String(condition.value ?? "").length)
            : 1 +
              String(cellValue ?? "").length +
              String(condition.value ?? "").length),
        comparisonWork > filterEvaluationLimits.comparisonWork)
      )
        throw Error(
          "Filter comparison work exceeds its 5,000,000-unit budget. Narrow the range or simplify criteria.",
        );
      return matchesTextFilter(cellValue, condition, workbook, isCaseSensitive);
    },
    fields = Array.from(
      { length: rangeAddress.lastCol - rangeAddress.firstCol + 1 },
      (O, fieldOffset) => ({
        column: rangeAddress.firstCol + fieldOffset,
        label:
          String(
            assertCalculatedBeforeFilter(
              workbookWithSettings,
              sheetId,
              `${columnIndexToLetters(rangeAddress.firstCol + fieldOffset)}${rangeAddress.firstRow + 1}`,
              workbook,
            ) ?? "",
          ) || columnIndexToLetters(rangeAddress.firstCol + fieldOffset),
      }),
    );
  for (
    let rowIndex = rangeAddress.firstRow + 1;
    rowIndex <= rangeAddress.lastRow;
    rowIndex++
  ) {
    const rowValues = fields.map((field) =>
        assertCalculatedBeforeFilter(
          workbookWithSettings,
          sheetId,
          `${columnIndexToLetters(field.column)}${rowIndex + 1}`,
          workbook,
        ),
      ),
      valueAtColumn = (sheetColumnIndex) =>
        rowValues[sheetColumnIndex - rangeAddress.firstCol];
    if (
      ((comparisonWork += rowValues.reduce(
        (total, rowCellValue) =>
          total +
          1 +
          (typeof rowCellValue == "string" ? rowCellValue.length : 0),
        0,
      )),
      comparisonWork > filterEvaluationLimits.comparisonWork)
    )
      throw Error(
        "Filter comparison work exceeds its 5,000,000-unit budget. Narrow the range or simplify criteria.",
      );
    (sourceRows.push(rowIndex + 1),
      sheetOrThrow.rowOriginIds?.[rowIndex] &&
        (originalSourceIds[rowIndex + 1] =
          sheetOrThrow.rowOriginIds[rowIndex]));
    const isRowHidden = axisVisibilityIndex.isHidden(rowIndex);
    isRowHidden && manuallyHiddenRows++;
    let isMatchingRow =
      definition.columns.every(
        (column, columnPosition) =>
          (columnValueSets[columnPosition] === null ||
            columnValueSets[columnPosition].has(
              comparableCellValueKey(
                valueAtColumn(column.column),
                isCaseSensitive,
              ),
            )) &&
          (!column.conditions?.length ||
            (column.join === "or"
              ? column.conditions.some((condition) =>
                  evaluateCondition(valueAtColumn(column.column), condition),
                )
              : column.conditions.every((condition) =>
                  evaluateCondition(valueAtColumn(column.column), condition),
                ))),
      ) &&
      (!definition.advanced?.length ||
        definition.advanced.some((conditionGroup) =>
          conditionGroup.every((groupCondition) =>
            evaluateCondition(
              valueAtColumn(groupCondition.column),
              groupCondition,
            ),
          ),
        ));
    if (isMatchingRow && definition.unique) {
      const uniqueKey = JSON.stringify(
        rowValues.map((uniqueCellValue) =>
          comparableCellValueKey(uniqueCellValue, isCaseSensitive),
        ),
      );
      seenUniqueRows.has(uniqueKey)
        ? (isMatchingRow = false)
        : seenUniqueRows.add(uniqueKey);
    }
    if (isMatchingRow)
      (matchingSourceRows.push(rowIndex + 1),
        isRowHidden
          ? hiddenMatchingRows++
          : visibleSourceRows.push(rowIndex + 1));
    else {
      const lastExcludedRange = excludedRows.at(-1);
      lastExcludedRange?.end === rowIndex - 1
        ? (lastExcludedRange.end = rowIndex)
        : excludedRows.push({ start: rowIndex, end: rowIndex });
    }
  }
  const hiddenIntervals = [];
  for (const interval of [...cn(sheetOrThrow, "row"), ...excludedRows].sort(
    (leftInterval, rightInterval) => leftInterval.start - rightInterval.start,
  )) {
    const lastInterval = hiddenIntervals.at(-1);
    lastInterval && interval.start <= lastInterval.end + 1
      ? (lastInterval.end = Math.max(lastInterval.end, interval.end))
      : hiddenIntervals.push({ ...interval });
  }
  return {
    hiddenRows: normalizeHiddenIntervals(
      hiddenIntervals,
      sheetOrThrow.rowCount,
    ),
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
function buildFilterFacetValues(
  workbook,
  sheetId,
  filterRequest,
  facetColumn,
  searchText = "",
) {
  const definition = validateFilterRequest(workbook, sheetId, filterRequest),
    rangeAddress = parseRangeAddress(definition.range);
  if (
    !Number.isInteger(facetColumn) ||
    facetColumn < rangeAddress.firstCol ||
    facetColumn > rangeAddress.lastCol
  )
    throw Error("Facet column is outside the filter range.");
  const workbookWithSettings = calculateWorkbookWithSettings(workbook),
    filterResult = applyFilterConditions(
      workbook,
      sheetId,
      {
        ...definition,
        unique: false,
        columns: definition.columns.filter(
          (column) => column.column !== facetColumn,
        ),
      },
      workbookWithSettings,
    ),
    valueCounts = new Map();
  for (const matchingSourceRow of filterResult.context.matchingSourceRows) {
    const facetValue = assertCalculatedBeforeFilter(
        workbookWithSettings,
        sheetId,
        `${columnIndexToLetters(facetColumn)}${matchingSourceRow}`,
        workbook,
      ),
      valueKey = comparableCellValueKey(facetValue, definition.caseSensitive);
    if (
      !normalizeComparableText(String(facetValue ?? "(Blanks)")).includes(
        normalizeComparableText(searchText),
      )
    )
      continue;
    const existingEntry = valueCounts.get(valueKey);
    existingEntry
      ? existingEntry.count++
      : valueCounts.set(valueKey, {
          value: facetValue,
          count: 1,
          key: valueKey,
        });
  }
  return {
    values: [...valueCounts.values()].slice(0, filterEvaluationLimits.values),
    distinct: valueCounts.size,
    truncated: valueCounts.size > filterEvaluationLimits.values,
  };
}
function validateSavedFilterViews(savedFilterViews, targetWorkbook) {
  if (savedFilterViews === undefined) return;
  if (
    !Array.isArray(savedFilterViews) ||
    savedFilterViews.length > savedFilterViewLimits.views ||
    JSON.stringify(savedFilterViews).length > savedFilterViewLimits.characters
  )
    throw Error(
      "Saved filter views are limited to 20 views and 512,000 metadata characters per workbook.",
    );
  const seenIds = new Set(),
    seenNames = new Set();
  targetWorkbook.savedFilterViews = savedFilterViews.map((savedFilterView) => {
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
      seenIds.has(savedFilterView.id) ||
      seenNames.has(normalizeNameForComparison(savedFilterViewName))
    )
      throw Error("Saved filter view names and IDs must be unique.");
    (seenIds.add(savedFilterView.id),
      seenNames.add(normalizeNameForComparison(savedFilterViewName)),
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
    const validateConditionMetadata = (conditionMetadata) =>
      validateSavedFilterViewMetadata(conditionMetadata, [
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
        column.conditions.forEach(validateConditionMetadata);
      }
    if (savedFilterView.definition.advanced !== undefined) {
      if (!Array.isArray(savedFilterView.definition.advanced))
        throw Error("Invalid saved advanced criteria.");
      for (const advancedGroup of savedFilterView.definition.advanced) {
        if (!Array.isArray(advancedGroup))
          throw Error("Invalid saved advanced criteria.");
        advancedGroup.forEach(validateConditionMetadata);
      }
    }
    const workbook = {
        ...targetWorkbook,
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
        (populationCount) =>
          !Number.isInteger(populationCount) || Number(populationCount) < 0,
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
  const copiedWorkbook = { ...workbook };
  return (
    validateSavedFilterViews(savedFilterViews, copiedWorkbook),
    copiedWorkbook
  );
}
function createSavedFilterView(
  workbook,
  sheetId,
  definition,
  viewName,
  viewId,
) {
  const sheetOrThrow = findSheetOrThrow(workbook, sheetId);
  if (
    (assertSavedFilterViewsEditable(workbook, sheetId),
    viewId && findSavedFilterView(workbook, viewId).sheetId !== sheetId)
  )
    throw Error("A saved view must keep its source worksheet.");
  const context = applyFilterConditions(workbook, sheetId, definition).context,
    savedFilterView = {
      id: viewId ?? crypto.randomUUID(),
      name: normalizeSavedFilterViewName(viewName),
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
    ...(workbook.savedFilterViews ?? []).filter(
      (existingView) => existingView.id !== viewId,
    ),
    savedFilterView,
  ]);
}
function addSavedFilterView(workbook, viewId, newName) {
  return (
    assertSavedFilterViewsEditable(
      workbook,
      findSavedFilterView(workbook, viewId).sheetId,
    ),
    copyWithSavedFilterViews(
      workbook,
      workbook.savedFilterViews.map((savedFilterView) =>
        savedFilterView.id === viewId
          ? { ...savedFilterView, name: normalizeSavedFilterViewName(newName) }
          : savedFilterView,
      ),
    )
  );
}
function removeSavedFilterView(workbook, viewId) {
  return (
    assertSavedFilterViewsEditable(
      workbook,
      findSavedFilterView(workbook, viewId).sheetId,
    ),
    copyWithSavedFilterViews(
      workbook,
      workbook.savedFilterViews.filter(
        (savedFilterView) => savedFilterView.id !== viewId,
      ),
    )
  );
}
function applyPrintLayoutChange(workbook, sheetId, layout, options = {}) {
  if (options.readOnly) throw Error("This workbook is read-only.");
  options.expectedSnapshot &&
    verifyPrintSnapshotGuard(workbook, options.expectedSnapshot);
  const sheetOrThrow = findSheetOrThrow(workbook, sheetId);
  assertSheetUnprotected(sheetOrThrow);
  const updatedSheet = { ...sheetOrThrow };
  return (
    validatePageSetup(layout, updatedSheet),
    zU(updatedSheet, updatedSheet.pageLayout),
    JSON.stringify(updatedSheet.pageLayout) ===
    JSON.stringify(sheetOrThrow.pageLayout)
      ? workbook
      : {
          ...workbook,
          sheets: workbook.sheets.map((sheet) =>
            sheet.id === sheetId ? updatedSheet : sheet,
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
