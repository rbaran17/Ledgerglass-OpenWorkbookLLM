import { isFormulaInputValue } from "./11-local-workbook-store-and-versions.js";
import { eI, formatCellAddress } from "./04-cell-input-paste-and-fill.js";
import { throwWithMessage, isFormulaCellValue } from "./03-workbook-model.js";
import { coerceCellValueToText } from "./16-transformation-recipe-engine.js";
import { X } from "lucide-react";
import i from "react/jsx-runtime";
import ne from "react";
import { workbookLimits } from "./02-text-and-financial-functions.js";
import {
  formatNumberWithPattern,
  findSheetOrThrow,
  formatOccupiedRangeAddress,
  calculateWorkbookWithSettings,
  _t,
  cn,
  parseRangeAddress,
  parseSelectionAreas,
  normalizeCellRecordInput,
  createAxisVisibilityIndex,
  columnIndexToLetters,
  rewriteStructuredReferences,
} from "./03-workbook-model.js";
import {
  useWorkbookMessages,
  LocalizedText,
} from "./04-cell-input-paste-and-fill.js";
import { importWorkbookFile } from "./07-file-import-and-worker-dispatch.js";
import { useCancellableOperation } from "./09-workbook-links.js";
import {
  patternCheckBudgets,
  Bj,
  isBlankFormulaCell,
  computeFormulaSignature,
  identifierPattern,
  sheetSpanPattern,
  isNonReferenceToken,
  definedNameReferencePattern,
  referenceTokenPattern,
  ww,
  readFormulaText,
} from "./51-pattern-checks-and-citations.js";

function runWorkbookComparison(
  before,
  workbook,
  comparisonRequest,
  options = {},
) {
  if (
    !comparisonRequest ||
    !Array.isArray(comparisonRequest.mappings) ||
    !comparisonRequest.mappings.length ||
    comparisonRequest.mappings.length > patternCheckBudgets.mappings
  )
    throw Error("Choose 1–20 explicit worksheet/range mappings.");
  if (
    options.includeHidden !== undefined &&
    typeof options.includeHidden != "boolean"
  )
    throw Error("Include hidden must be boolean.");
  const allowedSheetIdsFor = (workbook, allowedIds) => {
      if (
        allowedIds !== undefined &&
        (!Array.isArray(allowedIds) ||
          allowedIds.some(
            (candidateId) =>
              typeof candidateId != "string" ||
              !workbook.sheets.some((sheet) => sheet.id === candidateId),
          ))
      )
        throw Error("The comparison worksheet access context is invalid.");
      return [
        ...new Set(allowedIds ?? workbook.sheets.map((sheet) => sheet.id)),
      ].sort();
    },
    allowedSheetIds = {
      before: allowedSheetIdsFor(before, options.beforeAllowedSheetIds),
      after: allowedSheetIdsFor(workbook, options.afterAllowedSheetIds),
    },
    workbooksBySide = { before: before, after: workbook },
    report = {
      before: { id: before.id, name: before.name, revision: before.revision },
      after: {
        id: workbook.id,
        name: workbook.name,
        revision: workbook.revision,
      },
      request: structuredClone(comparisonRequest),
      context: {
        beforeAllowedSheetIds: allowedSheetIds.before,
        afterAllowedSheetIds: allowedSheetIds.after,
        includeHidden: !!options.includeHidden,
      },
      findings: [],
      counts: {
        beforeRows: 0,
        afterRows: 0,
        matchedRows: 0,
        insertedRows: 0,
        deletedRows: 0,
        unresolvedBeforeRows: 0,
        unresolvedAfterRows: 0,
        scannedCells: 0,
        scannedFormulas: 0,
      },
      coverage: {
        selectedCells: 0,
        notes: [
          "Static formula text and stored cell metadata are compared; calculated values and imported caches are never used.",
          "Row keys use exact typed literal values; case and whitespace are significant. Duplicate, blank or formula keys remain unresolved.",
          "Literal input inventory is observed; formula-column hardcodes and pattern outliers are review candidates, not proof of error. Pattern analysis uses relative A1 signatures within each selected column; a dominant pattern requires at least two formulas and more than half the formulas.",
          "Only explicitly mapped rectangles receive row/cell comparison. Static links and hidden structure are inspected across permitted worksheets; unmapped sheets are listed.",
        ],
        formulaValuesCalculated: false,
        beforeUnmappedSheets: 0,
        afterUnmappedSheets: 0,
      },
    },
    addFinding = (finding) => {
      if (report.findings.length >= patternCheckBudgets.findings)
        throw Error(
          "Comparison exceeds 30,000 findings; narrow the source workbooks.",
        );
      report.findings.push(finding);
    },
    createLocation = (side, worksheet, address) => ({
      side: side,
      workbookId: workbooksBySide[side].id,
      revision: workbooksBySide[side].revision,
      sheetId: worksheet.id,
      sheetName: worksheet.name,
      address: address,
    }),
    createSideLocation = (sideName, targetSheet, locationAddress) => ({
      [sideName]: createLocation(sideName, targetSheet, locationAddress),
    }),
    hiddenSpansOf = (spanSheet) => ({
      rows: cn(spanSheet, "row"),
      cols: cn(spanSheet, "column"),
    }),
    readMappedRange = (rangeSide, rangeReference, hasHeaders) => {
      if (
        !rangeReference ||
        typeof rangeReference.sheetId != "string" ||
        typeof rangeReference.range != "string"
      )
        throw Error(
          "Choose explicit worksheet and rectangular range references.",
        );
      const sheet = workbooksBySide[rangeSide].sheets.find(
        (sheet) => sheet.id === rangeReference.sheetId,
      );
      if (!sheet || !allowedSheetIds[rangeSide].includes(sheet.id))
        throw Error(
          "A selected source is outside the authorized comparison context.",
        );
      const rangeAddress = parseRangeAddress(rangeReference.range),
        rowCount = rangeAddress.lastRow - rangeAddress.firstRow + 1,
        columnCount = rangeAddress.lastCol - rangeAddress.firstCol + 1;
      if (
        rowCount > patternCheckBudgets.rowsPerRange ||
        columnCount > patternCheckBudgets.columnsPerRange ||
        rangeAddress.lastRow >= sheet.rowCount ||
        rangeAddress.lastCol >= sheet.colCount
      )
        throw Error(
          "Each comparison range supports at most 1,000 rows and 100 columns within the worksheet.",
        );
      if (
        ((report.coverage.selectedCells += rowCount * columnCount),
        report.coverage.selectedCells > patternCheckBudgets.selectedCells)
      )
        throw Error(
          "Comparison supports at most 10,000 selected cells across both workbooks.",
        );
      const hiddenSpans = hiddenSpansOf(sheet);
      if (
        !options.includeHidden &&
        (sheet.hidden ||
          sheet.veryHidden ||
          hiddenSpans.rows.some(
            (row) =>
              row.start <= rangeAddress.lastRow &&
              row.end >= rangeAddress.firstRow,
          ) ||
          hiddenSpans.cols.some(
            (col) =>
              col.start <= rangeAddress.lastCol &&
              col.end >= rangeAddress.firstCol,
          ))
      )
        throw Error(
          "A selected range contains hidden content. Explicitly include permitted hidden content or choose a visible range.",
        );
      const headers = Array.from({ length: columnCount }, (j, columnOffset) => {
        if (!hasHeaders) return columnIndexToLetters(columnOffset);
        const cell =
          sheet.cells[
            formatCellAddress(
              rangeAddress.firstRow,
              rangeAddress.firstCol + columnOffset,
            )
          ];
        if (
          isFormulaInputValue(cell) ||
          typeof cell?.raw != "string" ||
          !cell.raw.trim()
        )
          throw Error(
            "Header rows require nonempty literal text in every column. Turn off headers for an unstructured sheet.",
          );
        return cell.raw;
      });
      if (
        new Set(headers.map((header) => header.toLowerCase())).size !==
        headers.length
      )
        throw Error(
          "Duplicate headers are ambiguous; resolve them or compare without headers.",
        );
      return {
        sheet: sheet,
        range: `${formatCellAddress(rangeAddress.firstRow, rangeAddress.firstCol)}:${formatCellAddress(rangeAddress.lastRow, rangeAddress.lastCol)}`,
        headers: headers,
        firstCol: rangeAddress.firstCol,
        rows: Array.from(
          { length: rowCount - Number(hasHeaders) },
          (j, rowOffset) => {
            const row = rangeAddress.firstRow + Number(hasHeaders) + rowOffset;
            return {
              row: row,
              cells: headers.map(
                (header, headerIndex) =>
                  sheet.cells[
                    formatCellAddress(row, rangeAddress.firstCol + headerIndex)
                  ],
              ),
            };
          },
        ),
      };
    },
    mappedSheetIds = { before: new Set(), after: new Set() },
    mappedRangesBySheet = { before: new Map(), after: new Map() };
  for (const mapping of comparisonRequest.mappings) {
    if (
      !mapping ||
      typeof mapping.headers != "boolean" ||
      !["keys", "position"].includes(mapping.rowMatch)
    )
      throw Error(
        "Each mapping requires explicit header and row-matching rules.",
      );
    for (const mappingSide of ["before", "after"]) {
      const sideMapping = mapping[mappingSide],
        rangeAddress = parseRangeAddress(sideMapping.range),
        existingRanges =
          mappedRangesBySheet[mappingSide].get(sideMapping.sheetId) ?? [];
      if (
        existingRanges.some(
          (existingRange) =>
            existingRange.firstRow <= rangeAddress.lastRow &&
            existingRange.lastRow >= rangeAddress.firstRow &&
            existingRange.firstCol <= rangeAddress.lastCol &&
            existingRange.lastCol >= rangeAddress.firstCol,
        )
      )
        throw Error("Comparison mappings may not overlap on either side.");
      (mappedRangesBySheet[mappingSide].set(sideMapping.sheetId, [
        ...existingRanges,
        rangeAddress,
      ]),
        mappedSheetIds[mappingSide].add(sideMapping.sheetId));
    }
    const beforeRange = readMappedRange(
        "before",
        mapping.before,
        mapping.headers,
      ),
      afterRange = readMappedRange("after", mapping.after, mapping.headers);
    ((report.counts.beforeRows += beforeRange.rows.length),
      (report.counts.afterRows += afterRange.rows.length));
    const rowLocation = (locationSide, mappedRange, rowNumber) =>
        createSideLocation(
          locationSide,
          mappedRange.sheet,
          `${formatCellAddress(rowNumber, mappedRange.firstCol)}:${formatCellAddress(rowNumber, mappedRange.firstCol + mappedRange.headers.length - 1)}`,
        ),
      matchedPairs = [],
      reportUnmatchedRow = (
        unmatchedSide,
        unmatchedRange,
        unmatchedRowIndex,
      ) => {
        (report.counts[
          unmatchedSide === "before" ? "deletedRows" : "insertedRows"
        ]++,
          addFinding({
            kind: unmatchedSide === "before" ? "deleted-row" : "inserted-row",
            confidence: "observed",
            message: `${unmatchedSide === "before" ? "Deleted" : "Inserted"} record under the declared ${mapping.rowMatch} match rule.`,
            ...rowLocation(
              unmatchedSide,
              unmatchedRange,
              unmatchedRange.rows[unmatchedRowIndex].row,
            ),
          }));
      },
      reportUnresolvedRow = (
        unresolvedSide,
        unresolvedRange,
        unresolvedRowIndex,
        unresolvedMessage,
      ) => {
        (report.counts[
          unresolvedSide === "before"
            ? "unresolvedBeforeRows"
            : "unresolvedAfterRows"
        ]++,
          addFinding({
            kind: "uncertain",
            confidence: "unresolved",
            message: unresolvedMessage,
            ...rowLocation(
              unresolvedSide,
              unresolvedRange,
              unresolvedRange.rows[unresolvedRowIndex].row,
            ),
          }));
      };
    if (mapping.rowMatch === "position") {
      if (mapping.keys?.length)
        throw Error("Positional matching does not use keys.");
      addFinding({
        kind: "uncertain",
        confidence: "unresolved",
        message:
          "Positional alignment was explicitly chosen. Reordering cannot be distinguished from changed records; use unique literal keys for record identity.",
        ...createSideLocation("before", beforeRange.sheet, beforeRange.range),
        ...createSideLocation("after", afterRange.sheet, afterRange.range),
      });
      for (
        let rowPosition = 0;
        rowPosition < Math.max(beforeRange.rows.length, afterRange.rows.length);
        rowPosition++
      )
        rowPosition >= beforeRange.rows.length
          ? reportUnmatchedRow("after", afterRange, rowPosition)
          : rowPosition >= afterRange.rows.length
            ? reportUnmatchedRow("before", beforeRange, rowPosition)
            : matchedPairs.push([rowPosition, rowPosition]);
    } else {
      if (
        !Array.isArray(mapping.keys) ||
        !mapping.keys.length ||
        mapping.keys.length > patternCheckBudgets.keys
      )
        throw Error(
          "Key matching requires 1–8 explicit header pairs (A, B, … without headers).",
        );
      const keyColumnIndexesOf = (keyRange, keySide) => {
          const keyColumnIndexes = mapping.keys.map((key) =>
            keyRange.headers.indexOf(key[keySide]),
          );
          if (
            keyColumnIndexes.some((columnIndex) => columnIndex < 0) ||
            new Set(keyColumnIndexes).size !== keyColumnIndexes.length
          )
            throw Error(
              "Key headers must exist exactly once and may not be repeated.",
            );
          return keyColumnIndexes;
        },
        indexRowsByKey = (keyedRange, indexSide) => {
          const indexKeyColumns = keyColumnIndexesOf(keyedRange, indexSide),
            rowsByKey = new Map();
          return (
            keyedRange.rows.forEach((row, rowIndex) => {
              if (
                indexKeyColumns.some(
                  (keyColumn) =>
                    isBlankFormulaCell(row.cells[keyColumn]) ||
                    isFormulaInputValue(row.cells[keyColumn]),
                )
              ) {
                reportUnresolvedRow(
                  indexSide,
                  keyedRange,
                  rowIndex,
                  "Blank or formula key: record identity is unresolved; formulas are not evaluated.",
                );
                return;
              }
              const rowKey = JSON.stringify(
                  indexKeyColumns.map((keyPartColumn) => [
                    typeof readFormulaText(row.cells[keyPartColumn]),
                    readFormulaText(row.cells[keyPartColumn]),
                  ]),
                ),
                matchingRowIndexes = rowsByKey.get(rowKey) ?? [];
              (matchingRowIndexes.push(rowIndex),
                rowsByKey.set(rowKey, matchingRowIndexes));
            }),
            rowsByKey
          );
        },
        beforeRowsByKey = indexRowsByKey(beforeRange, "before"),
        afterRowsByKey = indexRowsByKey(afterRange, "after");
      for (const groupKey of new Set([
        ...beforeRowsByKey.keys(),
        ...afterRowsByKey.keys(),
      ])) {
        const beforeRowIndexes = beforeRowsByKey.get(groupKey) ?? [],
          afterRowIndexes = afterRowsByKey.get(groupKey) ?? [];
        if (beforeRowIndexes.length > 1 || afterRowIndexes.length > 1) {
          for (const beforeRowIndex of beforeRowIndexes)
            reportUnresolvedRow(
              "before",
              beforeRange,
              beforeRowIndex,
              `Duplicate key ambiguity: ${beforeRowIndexes.length} before / ${afterRowIndexes.length} after records; no automatic pairing.`,
            );
          for (const afterRowIndex of afterRowIndexes)
            reportUnresolvedRow(
              "after",
              afterRange,
              afterRowIndex,
              `Duplicate key ambiguity: ${beforeRowIndexes.length} before / ${afterRowIndexes.length} after records; no automatic pairing.`,
            );
        } else
          beforeRowIndexes.length && afterRowIndexes.length
            ? matchedPairs.push([beforeRowIndexes[0], afterRowIndexes[0]])
            : beforeRowIndexes.length
              ? reportUnmatchedRow("before", beforeRange, beforeRowIndexes[0])
              : reportUnmatchedRow("after", afterRange, afterRowIndexes[0]);
      }
    }
    for (const [matchedBeforeIndex, matchedAfterIndex] of matchedPairs) {
      report.counts.matchedRows++;
      const beforeRow = beforeRange.rows[matchedBeforeIndex],
        afterRow = afterRange.rows[matchedAfterIndex];
      mapping.rowMatch === "keys" &&
        beforeRow.row !== afterRow.row &&
        addFinding({
          kind: "moved-row",
          confidence: "observed",
          message: "The unique typed key matches at a different worksheet row.",
          ...rowLocation("before", beforeRange, beforeRow.row),
          ...rowLocation("after", afterRange, afterRow.row),
        });
      for (const headerName of new Set([
        ...beforeRange.headers,
        ...afterRange.headers,
      ])) {
        const beforeHeaderIndex = beforeRange.headers.indexOf(headerName),
          afterHeaderIndex = afterRange.headers.indexOf(headerName),
          beforeCell =
            beforeHeaderIndex < 0
              ? undefined
              : beforeRow.cells[beforeHeaderIndex],
          afterCell =
            afterHeaderIndex < 0 ? undefined : afterRow.cells[afterHeaderIndex];
        (beforeHeaderIndex >= 0 &&
          afterHeaderIndex >= 0 &&
          ww(beforeCell ?? { raw: null }, afterCell ?? { raw: null })) ||
          addFinding({
            kind:
              isFormulaInputValue(beforeCell) && !isFormulaInputValue(afterCell)
                ? "formula-to-value"
                : "cell-change",
            confidence: "observed",
            message: `${headerName}: ${beforeHeaderIndex < 0 ? "column added" : afterHeaderIndex < 0 ? "column removed" : isFormulaInputValue(beforeCell) && !isFormulaInputValue(afterCell) ? "formula replaced by a literal or blank" : "stored value, formula or cell metadata changed"}.`,
            ...(beforeHeaderIndex < 0
              ? {}
              : {
                  ...createSideLocation(
                    "before",
                    beforeRange.sheet,
                    formatCellAddress(
                      beforeRow.row,
                      beforeRange.firstCol + beforeHeaderIndex,
                    ),
                  ),
                  beforeValue: readFormulaText(beforeCell),
                }),
            ...(afterHeaderIndex < 0
              ? {}
              : {
                  ...createSideLocation(
                    "after",
                    afterRange.sheet,
                    formatCellAddress(
                      afterRow.row,
                      afterRange.firstCol + afterHeaderIndex,
                    ),
                  ),
                  afterValue: readFormulaText(afterCell),
                }),
          });
      }
    }
    for (const [reportSide, reportRange] of [
      ["before", beforeRange],
      ["after", afterRange],
    ])
      for (
        let headerPosition = 0;
        headerPosition < reportRange.headers.length;
        headerPosition++
      ) {
        const rowIndexesByPattern = new Map();
        let formulaCount = 0;
        reportRange.rows.forEach((row, formulaRowIndex) => {
          if (isFormulaInputValue(row.cells[headerPosition])) {
            formulaCount++;
            try {
              const formulaSignature = computeFormulaSignature(
                  row.cells[headerPosition].raw,
                  formatCellAddress(
                    row.row,
                    reportRange.firstCol + headerPosition,
                  ),
                ),
                signatureRowIndexes =
                  rowIndexesByPattern.get(formulaSignature) ?? [];
              (signatureRowIndexes.push(formulaRowIndex),
                rowIndexesByPattern.set(formulaSignature, signatureRowIndexes));
            } catch {
              addFinding({
                kind: "uncertain",
                confidence: "unresolved",
                message: "Formula pattern could not be normalized reliably.",
                ...createSideLocation(
                  reportSide,
                  reportRange.sheet,
                  formatCellAddress(
                    row.row,
                    reportRange.firstCol + headerPosition,
                  ),
                ),
              });
            }
          }
        });
        const patternEntries = [...rowIndexesByPattern.entries()].sort(
            (leftEntry, rightEntry) =>
              rightEntry[1].length - leftEntry[1].length,
          ),
          dominantPattern = patternEntries[0];
        for (const row of reportRange.rows) {
          const columnCell = row.cells[headerPosition];
          if (isFormulaInputValue(columnCell) || isBlankFormulaCell(columnCell))
            continue;
          const isFormulaDominant =
            formulaCount > reportRange.rows.length / 2 &&
            !!dominantPattern &&
            dominantPattern[1].length >= 2 &&
            dominantPattern[1].length > formulaCount / 2;
          addFinding({
            kind: "hardcode",
            confidence: isFormulaDominant ? "candidate" : "observed",
            message: isFormulaDominant
              ? `${reportRange.headers[headerPosition]}: literal value in a formula-dominant column; a deliberate exception is possible.`
              : `${reportRange.headers[headerPosition]}: stored literal input, not a formula. This inventory entry is not an error finding.`,
            ...createSideLocation(
              reportSide,
              reportRange.sheet,
              formatCellAddress(row.row, reportRange.firstCol + headerPosition),
            ),
            [reportSide === "before" ? "beforeValue" : "afterValue"]:
              readFormulaText(columnCell),
          });
        }
        if (!(formulaCount < 2)) {
          if (
            !dominantPattern ||
            dominantPattern[1].length < 2 ||
            dominantPattern[1].length <= formulaCount / 2
          ) {
            rowIndexesByPattern.size > 1 &&
              addFinding({
                kind: "uncertain",
                confidence: "unresolved",
                message: `${reportRange.headers[headerPosition]}: no majority formula pattern; ${rowIndexesByPattern.size} patterns among ${formulaCount} formulas. No anomaly winner was selected.`,
                ...createSideLocation(
                  reportSide,
                  reportRange.sheet,
                  reportRange.range,
                ),
              });
            continue;
          }
          reportRange.rows.forEach((row, columnRowIndex) => {
            const candidateCell = row.cells[headerPosition],
              cellAddress = formatCellAddress(
                row.row,
                reportRange.firstCol + headerPosition,
              );
            isFormulaInputValue(candidateCell) &&
              !dominantPattern[1].includes(columnRowIndex) &&
              addFinding({
                kind: "formula-pattern",
                confidence: "candidate",
                message: `${reportRange.headers[headerPosition]}: formula differs from the dominant relative-reference pattern (${dominantPattern[1].length}/${formulaCount} formulas); review against ${formatCellAddress(reportRange.rows[dominantPattern[1][0]].row, reportRange.firstCol + headerPosition)}.`,
                ...createSideLocation(
                  reportSide,
                  reportRange.sheet,
                  cellAddress,
                ),
                [reportSide === "before" ? "beforeValue" : "afterValue"]:
                  readFormulaText(candidateCell),
              });
          });
        }
      }
    (beforeRange.sheet.name !== afterRange.sheet.name ||
      !ww(beforeRange.headers, afterRange.headers)) &&
      addFinding({
        kind: "metadata",
        confidence: "observed",
        message:
          "Explicitly matched worksheet names, header text or column order differ.",
        ...createSideLocation("before", beforeRange.sheet, beforeRange.range),
        ...createSideLocation("after", afterRange.sheet, afterRange.range),
      });
    for (const propertyName of new Set([
      ...Object.keys(beforeRange.sheet),
      ...Object.keys(afterRange.sheet),
    ]))
      ["cells", "id", "name"].includes(propertyName) ||
        ww(beforeRange.sheet[propertyName], afterRange.sheet[propertyName]) ||
        addFinding({
          kind: "metadata",
          confidence: "observed",
          message: `Worksheet property changed: ${propertyName}. Property content is not evaluated.`,
          ...createSideLocation("before", beforeRange.sheet, beforeRange.range),
          ...createSideLocation("after", afterRange.sheet, afterRange.range),
        });
  }
  let formulaCharacterCount = 0;
  for (const workbookSide of ["before", "after"]) {
    const workbook = workbooksBySide[workbookSide];
    for (const sheet of workbook.sheets) {
      if (!allowedSheetIds[workbookSide].includes(sheet.id)) continue;
      const sheetHiddenSpans = hiddenSpansOf(sheet),
        isContentIncluded =
          options.includeHidden || (!sheet.hidden && !sheet.veryHidden);
      (sheet.hidden || sheet.veryHidden) &&
        addFinding({
          kind: "hidden-content",
          confidence: "observed",
          message: `${sheet.veryHidden ? "Very hidden" : "Hidden"} worksheet; ${isContentIncluded ? "content included by explicit permission" : "cell content withheld"}.`,
          ...createSideLocation(workbookSide, sheet, ""),
        });
      for (const [spanKind, spanList] of [
        ["row", sheetHiddenSpans.rows],
        ["column", sheetHiddenSpans.cols],
      ])
        for (const hiddenSpan of spanList)
          addFinding({
            kind: "hidden-content",
            confidence: "observed",
            message: `Hidden ${spanKind} interval, including collapsed outlines; ${options.includeHidden ? "content included" : "cell content withheld"}.`,
            ...createSideLocation(
              workbookSide,
              sheet,
              spanKind === "row"
                ? `${hiddenSpan.start + 1}:${hiddenSpan.end + 1}`
                : `${columnIndexToLetters(hiddenSpan.start)}:${columnIndexToLetters(hiddenSpan.end)}`,
            ),
          });
      if (
        (mappedSheetIds[workbookSide].has(sheet.id) ||
          (report.coverage[
            workbookSide === "before"
              ? "beforeUnmappedSheets"
              : "afterUnmappedSheets"
          ]++,
          addFinding({
            kind: "unmapped-sheet",
            confidence: "unresolved",
            message:
              "Worksheet has no explicit counterpart mapping. No insertion, deletion or identity match was inferred.",
            ...createSideLocation(workbookSide, sheet, ""),
          })),
        !!isContentIncluded)
      )
        for (const [storedAddress, cell] of Object.entries(sheet.cells)) {
          const parsedAddress = _t(storedAddress);
          if (
            !options.includeHidden &&
            (sheetHiddenSpans.rows.some(
              (row) =>
                row.start <= parsedAddress.row && row.end >= parsedAddress.row,
            ) ||
              sheetHiddenSpans.cols.some(
                (col) =>
                  col.start <= parsedAddress.col &&
                  col.end >= parsedAddress.col,
              ))
          )
            continue;
          if (++report.counts.scannedCells > patternCheckBudgets.scannedCells)
            throw Error(
              "Static comparison supports 20,000 permitted stored cells across both workbooks.",
            );
          if (!isFormulaInputValue(cell)) continue;
          report.counts.scannedFormulas++;
          const reportCellFinding = (findingKind, findingMessage) =>
              addFinding({
                kind: findingKind,
                confidence:
                  findingKind === "uncertain" ? "unresolved" : "observed",
                message: findingMessage,
                ...createSideLocation(workbookSide, sheet, storedAddress),
              }),
            inspectFormula = (
              formulaText,
              contextSheetId,
              expansionStack = [],
            ) => {
              if (
                ((formulaCharacterCount += formulaText.length),
                formulaCharacterCount > patternCheckBudgets.formulaCharacters)
              )
                throw Error(
                  "Static formula inspection exceeds the 2,000,000-character work budget.",
                );
              if (formulaText.length > 4096) {
                reportCellFinding(
                  "uncertain",
                  "Formula exceeds the 4,096-character inspection bound.",
                );
                return;
              }
              if (expansionStack.length >= patternCheckBudgets.nameDepth) {
                reportCellFinding(
                  "uncertain",
                  "Defined-name expansion exceeds depth 16.",
                );
                return;
              }
              let parsePosition = formulaText.startsWith("=") ? 1 : 0;
              const resolveReference = (referencedSheet, rangeText) => {
                  if (!referencedSheet) {
                    reportCellFinding(
                      "broken-link",
                      "Referenced worksheet is missing.",
                    );
                    return;
                  }
                  if (
                    !allowedSheetIds[workbookSide].includes(
                      referencedSheet.id,
                    ) ||
                    (!options.includeHidden &&
                      (referencedSheet.hidden || referencedSheet.veryHidden))
                  ) {
                    reportCellFinding(
                      "uncertain",
                      "A reference points outside the permitted visible worksheet context; target details are withheld.",
                    );
                    return;
                  }
                  try {
                    const normalizedRange = rangeText.replace(/[\s$]/g, ""),
                      rangeAddress = parseRangeAddress(
                        /^[A-Za-z]+:[A-Za-z]+$/.test(normalizedRange)
                          ? normalizedRange.replace(":", "1:") +
                              workbookLimits.maxRows
                          : /^\d+:\d+$/.test(normalizedRange)
                            ? "A" + normalizedRange.replace(":", ":XFD")
                            : normalizedRange,
                      );
                    if (
                      rangeAddress.lastRow >= referencedSheet.rowCount ||
                      rangeAddress.lastCol >= referencedSheet.colCount
                    )
                      throw Error();
                    const referencedHiddenSpans =
                      hiddenSpansOf(referencedSheet);
                    !options.includeHidden &&
                      (referencedHiddenSpans.rows.some(
                        (row) =>
                          row.start <= rangeAddress.lastRow &&
                          row.end >= rangeAddress.firstRow,
                      ) ||
                        referencedHiddenSpans.cols.some(
                          (col) =>
                            col.start <= rangeAddress.lastCol &&
                            col.end >= rangeAddress.firstCol,
                        )) &&
                      reportCellFinding(
                        "uncertain",
                        "A reference intersects hidden cells; their content was not inspected.",
                      );
                  } catch {
                    reportCellFinding(
                      "broken-link",
                      "Reference has invalid or out-of-bounds coordinates.",
                    );
                  }
                },
                resolveDefinedName = (definedNameText, scopeSheetId) => {
                  const matchedName =
                    workbook.definedNames?.find(
                      (definedName) =>
                        definedName.scopeSheetId === scopeSheetId &&
                        definedName.name.toLowerCase() ===
                          definedNameText.toLowerCase(),
                    ) ??
                    workbook.definedNames?.find(
                      (definedName) =>
                        !definedName.scopeSheetId &&
                        definedName.name.toLowerCase() ===
                          definedNameText.toLowerCase(),
                    );
                  if (!matchedName) {
                    reportCellFinding(
                      "broken-link",
                      "An unresolved defined-name reference is present.",
                    );
                    return;
                  }
                  const nameKey = `${matchedName.scopeSheetId ?? ""}:${matchedName.name.toLowerCase()}`;
                  if (expansionStack.includes(nameKey)) {
                    reportCellFinding(
                      "uncertain",
                      "Defined-name cycle: references cannot be completely inspected.",
                    );
                    return;
                  }
                  inspectFormula(
                    matchedName.refersTo,
                    matchedName.scopeSheetId ?? scopeSheetId,
                    [...expansionStack, nameKey],
                  );
                };
              for (; parsePosition < formulaText.length;) {
                const remainingFormula = formulaText.slice(parsePosition);
                if (remainingFormula[0] === '"') {
                  const stringLiteralMatch = /^"(?:[^"]|"")*"/.exec(
                    remainingFormula,
                  );
                  if (!stringLiteralMatch) {
                    reportCellFinding(
                      "uncertain",
                      "Unterminated formula text; link inspection is incomplete.",
                    );
                    break;
                  }
                  parsePosition += stringLiteralMatch[0].length;
                  continue;
                }
                const externalLinkMatch =
                  /^(?:'(?:[^']|'')*\[[^\]]+\](?:[^']|'')*'\s*!\s*[^+*/(),;\s]+|\[[^\]]+\][^!+*/(),;]*![^+*/(),;\s]+)/.exec(
                    remainingFormula,
                  );
                if (externalLinkMatch) {
                  (reportCellFinding(
                    "external-link",
                    "External workbook link is inventoried but never opened; target validity is unverified.",
                  ),
                    (parsePosition += externalLinkMatch[0].length));
                  continue;
                }
                if (
                  sheetSpanPattern.test(remainingFormula) ||
                  /^'(?:[^']|'')*:(?:[^']|'')*'\s*!/.test(remainingFormula)
                ) {
                  reportCellFinding(
                    "uncertain",
                    "Three-dimensional references require separate review and are not expanded.",
                  );
                  break;
                }
                const errorTokenMatch = /^#[A-Za-z0-9/?!]+/.exec(
                  remainingFormula,
                );
                if (errorTokenMatch) {
                  (errorTokenMatch[0].toUpperCase() === "#REF!" &&
                    reportCellFinding(
                      "broken-link",
                      "Stored #REF! token identifies a broken reference.",
                    ),
                    (parsePosition += errorTokenMatch[0].length));
                  continue;
                }
                const structuredRefMatch = /^([A-Za-z_][A-Za-z0-9_.]*)?\[/.exec(
                  remainingFormula,
                );
                if (structuredRefMatch) {
                  const structuredRefStart = parsePosition;
                  parsePosition += structuredRefMatch[0].length;
                  let bracketDepth = 1;
                  for (; parsePosition < formulaText.length && bracketDepth;)
                    (formulaText[parsePosition] === "[" && bracketDepth++,
                      formulaText[parsePosition] === "]" && bracketDepth--,
                      parsePosition++);
                  const structuredRefText = formulaText.slice(
                    structuredRefStart,
                    parsePosition,
                  );
                  try {
                    const structuredRef =
                      !bracketDepth &&
                      rewriteStructuredReferences(
                        workbook,
                        contextSheetId,
                        structuredRefText,
                        storedAddress,
                      );
                    structuredRef
                      ? resolveReference(
                          workbook.sheets.find(
                            (sheet) => sheet.id === structuredRef.sheetId,
                          ),
                          structuredRef.range,
                        )
                      : reportCellFinding(
                          "uncertain",
                          "Structured table reference could not be resolved; table/column or syntax may be unsupported.",
                        );
                  } catch {
                    reportCellFinding(
                      "uncertain",
                      "Structured reference could not be inspected.",
                    );
                  }
                  continue;
                }
                const functionCallMatch =
                  /^([A-Za-z_\\][A-Za-z0-9_.\\]*)\s*\(/.exec(remainingFormula);
                if (functionCallMatch) {
                  (/^(?:_xlfn\.)?(?:INDIRECT|OFFSET)$/i.test(
                    functionCallMatch[1],
                  ) &&
                    reportCellFinding(
                      "uncertain",
                      "Dynamic reference function is not executed; actual target is unverified.",
                    ),
                    /^(?:_xlfn\.)?(?:WEBSERVICE|RTD|CUBE\w*)$/i.test(
                      functionCallMatch[1],
                    ) &&
                      reportCellFinding(
                        "external-link",
                        "External/network-enabled function is inventoried but never executed.",
                      ),
                    (parsePosition += functionCallMatch[0].length));
                  continue;
                }
                const referenceMatch =
                  referenceTokenPattern.exec(remainingFormula);
                if (
                  referenceMatch &&
                  isNonReferenceToken(
                    remainingFormula[referenceMatch[0].length] ?? "",
                  )
                ) {
                  const referenceSheetName = referenceMatch[1]
                    ? Bj(referenceMatch[1])
                    : undefined;
                  (resolveReference(
                    referenceSheetName === undefined
                      ? workbook.sheets.find(
                          (sheet) => sheet.id === contextSheetId,
                        )
                      : workbook.sheets.find(
                          (sheet) =>
                            sheet.name.toLowerCase() ===
                            referenceSheetName.toLowerCase(),
                        ),
                    referenceMatch[2],
                  ),
                    referenceMatch[3] &&
                      reportCellFinding(
                        "uncertain",
                        "Spill reference is inspected at its anchor; runtime extent is not calculated.",
                      ),
                    (parsePosition += referenceMatch[0].length));
                  continue;
                }
                const definedNameMatch =
                  definedNameReferencePattern.exec(remainingFormula);
                if (definedNameMatch) {
                  const sheet = workbook.sheets.find(
                    (sheet) =>
                      sheet.name.toLowerCase() ===
                      Bj(definedNameMatch[1]).toLowerCase(),
                  );
                  (sheet
                    ? !allowedSheetIds[workbookSide].includes(sheet.id) ||
                      (!options.includeHidden &&
                        (sheet.hidden || sheet.veryHidden))
                      ? reportCellFinding(
                          "uncertain",
                          "Qualified name is outside the permitted context.",
                        )
                      : resolveDefinedName(definedNameMatch[2], sheet.id)
                    : reportCellFinding(
                        "broken-link",
                        "Defined-name worksheet qualifier is missing.",
                      ),
                    (parsePosition += definedNameMatch[0].length));
                  continue;
                }
                const numberMatch =
                  /^(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?/.exec(
                    remainingFormula,
                  );
                if (numberMatch) {
                  parsePosition += numberMatch[0].length;
                  continue;
                }
                const identifierMatch =
                  identifierPattern.exec(remainingFormula);
                if (identifierMatch) {
                  (/^(?:TRUE|FALSE)$/i.test(identifierMatch[0]) ||
                    (/^R(?:\d+|\[-?\d+\])C(?:\d+|\[-?\d+\])$/i.test(
                      identifierMatch[0],
                    )
                      ? reportCellFinding(
                          "uncertain",
                          "R1C1 syntax is not resolved.",
                        )
                      : resolveDefinedName(identifierMatch[0], contextSheetId)),
                    (parsePosition += identifierMatch[0].length));
                  continue;
                }
                (remainingFormula[0] === "|" &&
                  reportCellFinding(
                    "external-link",
                    "DDE-style syntax is inventoried and never executed.",
                  ),
                  parsePosition++);
              }
            };
          inspectFormula(cell.raw, sheet.id);
        }
    }
  }
  return (
    before.name !== workbook.name &&
      addFinding({
        kind: "metadata",
        confidence: "observed",
        message:
          "Workbook names differ; identities and revisions are retained separately.",
      }),
    structuredClone(report)
  );
}
const comparisonFindingKinds = [
    "all",
    "inserted-row",
    "deleted-row",
    "moved-row",
    "cell-change",
    "formula-to-value",
    "formula-pattern",
    "hardcode",
    "broken-link",
    "external-link",
    "uncertain",
    "hidden-content",
    "metadata",
    "unmapped-sheet",
  ],
  formatComparisonValue = (value) =>
    value === undefined
      ? "—"
      : value === null
        ? "(blank)"
        : typeof value == "boolean"
          ? String(value).toUpperCase()
          : String(value);
function WorkbookComparisonDialog(props) {
  const { text, language } = useWorkbookMessages(),
    cancellableOperation = useCancellableOperation(),
    dialogId = ne.useId(),
    dialogRef = ne.useRef(null),
    requestCounterRef = ne.useRef(0),
    mountedRef = ne.useRef(true),
    [importedWorkbook, setImportedWorkbook] = ne.useState(null),
    [importWarnings, setImportWarnings] = ne.useState([]),
    [isReadingBaseline, setIsReadingBaseline] = ne.useState(false),
    [isHiddenRequested, setIsHiddenRequested] = ne.useState(false),
    [mappings, setMappings] = ne.useState([]),
    [comparisonResult, setComparisonResult] = ne.useState(null),
    [comparedStateKey, setComparedStateKey] = ne.useState(""),
    [errorMessage, setErrorMessage] = ne.useState(""),
    [errorStateKey, setErrorStateKey] = ne.useState(""),
    [findingKindFilter, setFindingKindFilter] = ne.useState("all"),
    [pageIndex, setPageIndex] = ne.useState(0),
    workbook = importedWorkbook ?? props.baseline,
    afterAllowedSheetIds = [
      ...new Set(
        props.allowedSheetIds ?? props.workbook.sheets.map((sheet) => sheet.id),
      ),
    ].sort(),
    beforeAllowedSheetIds = [
      ...new Set(
        (importedWorkbook ? undefined : props.baselineAllowedSheetIds) ??
          workbook?.sheets.map((sheet) => sheet.id) ??
          [],
      ),
    ].sort(),
    isHiddenIncluded = !!props.allowHidden && isHiddenRequested,
    afterSheets = props.workbook.sheets.filter(
      (sheet) =>
        afterAllowedSheetIds.includes(sheet.id) &&
        (isHiddenIncluded || (!sheet.hidden && !sheet.veryHidden)),
    ),
    beforeSheets =
      workbook?.sheets.filter(
        (sheet) =>
          beforeAllowedSheetIds.includes(sheet.id) &&
          (isHiddenIncluded || (!sheet.hidden && !sheet.veryHidden)),
      ) ?? [],
    inputKey = JSON.stringify([
      props.workbook,
      props.sheetId,
      props.range,
      props.readOnly,
      props.allowedSheetIds,
      props.allowHidden,
      props.baseline,
      props.baselineAllowedSheetIds,
    ]),
    latestInputKeyRef = ne.useRef(inputKey);
  latestInputKeyRef.current = inputKey;
  const stateKey = JSON.stringify([
      inputKey,
      workbook,
      mappings,
      isHiddenIncluded,
      beforeAllowedSheetIds,
      afterAllowedSheetIds,
    ]),
    isResultCurrent = comparisonResult && stateKey === comparedStateKey,
    clearResult = () => {
      (setComparisonResult(null), setErrorMessage(""), setPageIndex(0));
    },
    createDefaultMapping = (
      workbook,
      isImportedBaseline = !!importedWorkbook,
    ) => {
      const sheet = workbook.sheets.find(
        (sheet) =>
          (isImportedBaseline ||
            !props.baselineAllowedSheetIds ||
            props.baselineAllowedSheetIds.includes(sheet.id)) &&
          (isHiddenIncluded || (!sheet.hidden && !sheet.veryHidden)),
      );
      return {
        before: {
          sheetId: sheet?.id ?? "",
          range: sheet ? formatOccupiedRangeAddress(sheet) : "A1",
        },
        after: { sheetId: props.sheetId, range: props.range },
        headers: true,
        rowMatch: "keys",
        keys: [{ before: "", after: "" }],
      };
    };
  (ne.useEffect(() => {
    const activeElement = document.activeElement;
    return (
      dialogRef.current
        ?.querySelector('[data-ui-label="Comparison baseline file"]')
        ?.focus(),
      (mountedRef.current = true),
      () => {
        ((mountedRef.current = false),
          requestCounterRef.current++,
          activeElement?.isConnected && activeElement.focus());
      }
    );
  }, []),
    ne.useEffect(() => {
      workbook &&
        !mappings.length &&
        setMappings([createDefaultMapping(workbook)]);
    }, [workbook]));
  const X$1 = (targetIndex, nextMapping) => {
      (setMappings(
        mappings.map((mapping, currentIndex) =>
          currentIndex === targetIndex ? nextMapping : mapping,
        ),
      ),
        clearResult());
    },
    importBaseline = async (file) => {
      if (!file) return;
      const requestNumber = ++requestCounterRef.current,
        startInputKey = inputKey;
      (setIsReadingBaseline(true),
        setErrorMessage(""),
        setComparisonResult(null),
        setImportedWorkbook(null),
        setImportWarnings([]),
        setMappings([]));
      try {
        if (!/\.(json|xlsx)$/i.test(file.name))
          throw Error("Choose a native JSON or XLSX comparison baseline.");
        const importResult = await importWorkbookFile(
          file,
          cancellableOperation.begin(),
        );
        if (
          !mountedRef.current ||
          requestNumber !== requestCounterRef.current ||
          latestInputKeyRef.current !== startInputKey
        )
          return;
        (setImportedWorkbook(importResult.workbook),
          setImportWarnings(importResult.warnings),
          setMappings([createDefaultMapping(importResult.workbook, true)]));
      } catch (error) {
        mountedRef.current &&
          requestNumber === requestCounterRef.current &&
          latestInputKeyRef.current === startInputKey &&
          (setErrorMessage(
            error instanceof Error ? error.message : "Baseline import failed.",
          ),
          setErrorStateKey(startInputKey));
      } finally {
        mountedRef.current &&
          requestNumber === requestCounterRef.current &&
          setIsReadingBaseline(false);
      }
    },
    onClick = () => {
      if (workbook)
        try {
          (setComparisonResult(
            runWorkbookComparison(
              workbook,
              props.workbook,
              { mappings: mappings },
              {
                beforeAllowedSheetIds: beforeAllowedSheetIds,
                afterAllowedSheetIds: afterAllowedSheetIds,
                includeHidden: isHiddenIncluded,
              },
            ),
          ),
            setComparedStateKey(stateKey),
            setPageIndex(0),
            setFindingKindFilter("all"),
            setErrorMessage(""));
        } catch (error) {
          (setComparisonResult(null),
            setErrorMessage(
              error instanceof Error
                ? error.message
                : "Comparison could not complete.",
            ),
            setErrorStateKey(inputKey));
        }
    },
    visibleFindings = isResultCurrent
      ? comparisonResult.findings.filter(
          (finding) =>
            findingKindFilter === "all" || finding.kind === findingKindFilter,
        )
      : [],
    pageCount = Math.max(1, Math.ceil(visibleFindings.length / 25)),
    currentPage = Math.min(pageIndex, pageCount - 1),
    closeDialog = () => {
      (requestCounterRef.current++, props.onClose());
    };
  return i.jsx("div", {
    className: "wc-backdrop",
    onMouseDown: (event) => {
      event.target === event.currentTarget && closeDialog();
    },
    children: i.jsxs("section", {
      ref: dialogRef,
      className: "wc-dialog",
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": `${dialogId}-title`,
      onKeyDown: (event) => {
        var lastFocusable, firstFocusable;
        if (
          (event.stopPropagation(),
          event.key === "Escape" && (event.preventDefault(), closeDialog()),
          event.key === "Tab")
        ) {
          const focusableElements = [
            ...(dialogRef.current?.querySelectorAll(
              'button:not(:disabled),input:not(:disabled),select:not(:disabled),textarea:not(:disabled),[tabindex="0"]',
            ) ?? []),
          ].filter((element) => !element.hidden);
          event.shiftKey && document.activeElement === focusableElements[0]
            ? (event.preventDefault(),
              (lastFocusable = focusableElements.at(-1)) == null ||
                lastFocusable.focus())
            : !event.shiftKey &&
              document.activeElement === focusableElements.at(-1) &&
              (event.preventDefault(),
              (firstFocusable = focusableElements[0]) == null ||
                firstFocusable.focus());
        }
      },
      children: [
        i.jsxs("header", {
          children: [
            i.jsxs("div", {
              children: [
                i.jsx("h2", {
                  id: `${dialogId}-title`,
                  children: i.jsx(LocalizedText, {
                    children: "Compare workbooks",
                  }),
                }),
                i.jsx("p", {
                  children: i.jsx(LocalizedText, {
                    children:
                      "Review record changes, formula consistency, links and hidden content.",
                  }),
                }),
              ],
            }),
            i.jsx("button", {
              "aria-label": text("Close workbook comparison"),
              onClick: closeDialog,
              lang: language("Close workbook comparison"),
              "data-ui-label": "Close workbook comparison",
              children: i.jsx(X, { size: 18 }),
            }),
          ],
        }),
        i.jsxs("div", {
          className: "wc-body",
          children: [
            i.jsx("p", {
              children: i.jsx(LocalizedText, {
                children:
                  "Comparison is read-only. Formulas, external links and imported cached values are never calculated. Both source workbooks stay unchanged.",
              }),
            }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Comparison baseline file" }),
                i.jsx("input", {
                  "aria-label": text("Comparison baseline file"),
                  type: "file",
                  accept: ".json,.xlsx",
                  onChange: (event) => {
                    (importBaseline(event.target.files?.[0]),
                      (event.target.value = ""));
                  },
                  lang: language("Comparison baseline file"),
                  "data-ui-label": "Comparison baseline file",
                }),
              ],
            }),
            isReadingBaseline &&
              i.jsx("p", {
                role: "status",
                children: i.jsx(LocalizedText, {
                  children: "Reading the comparison baseline…",
                }),
              }),
            workbook &&
              i.jsxs("p", {
                children: [
                  i.jsx(LocalizedText, { children: "Before: " }),
                  i.jsx("strong", { children: workbook.name }),
                  " ",
                  i.jsx(LocalizedText, { children: "· revision" }),
                  " ",
                  workbook.revision,
                  i.jsx(LocalizedText, { children: ". After: " }),
                  i.jsx("strong", { children: props.workbook.name }),
                  " ",
                  i.jsx(LocalizedText, { children: "· revision " }),
                  props.workbook.revision,
                  ".",
                ],
              }),
            importWarnings.length > 0 &&
              i.jsxs("details", {
                children: [
                  i.jsxs("summary", {
                    children: [
                      i.jsx(LocalizedText, {
                        children: "Baseline import notices (",
                      }),
                      importWarnings.length,
                      ")",
                    ],
                  }),
                  i.jsx("ul", {
                    children: importWarnings.map((warning, warningIndex) =>
                      i.jsx("li", { children: warning }, warningIndex),
                    ),
                  }),
                ],
              }),
            i.jsxs("label", {
              children: [
                i.jsx("input", {
                  type: "checkbox",
                  checked: isHiddenIncluded,
                  disabled: !props.allowHidden,
                  onChange: (event) => {
                    (setIsHiddenRequested(event.target.checked), clearResult());
                  },
                }),
                i.jsx(LocalizedText, {
                  children: "Include permitted hidden content",
                }),
              ],
            }),
            i.jsx("p", {
              children: i.jsx(LocalizedText, {
                children:
                  "Hidden structure is inventoried. Its cell content requires explicit permission. Imported files are separate user-selected baselines; host worksheet restrictions apply independently.",
              }),
            }),
            workbook &&
              mappings.map((mapping, mappingIndex) =>
                i.jsxs(
                  "fieldset",
                  {
                    children: [
                      i.jsxs("legend", {
                        children: [
                          i.jsx(LocalizedText, {
                            children: "Worksheet mapping ",
                          }),
                          mappingIndex + 1,
                        ],
                      }),
                      i.jsxs("div", {
                        className: "wc-fields",
                        children: [
                          i.jsxs("label", {
                            children: [
                              i.jsx(LocalizedText, {
                                children: "Before worksheet ",
                              }),
                              mappingIndex + 1,
                              i.jsxs("select", {
                                value: mapping.before.sheetId,
                                onChange: (event) =>
                                  X$1(mappingIndex, {
                                    ...mapping,
                                    before: {
                                      ...mapping.before,
                                      sheetId: event.target.value,
                                    },
                                  }),
                                children: [
                                  i.jsx("option", {
                                    value: "",
                                    children: "Choose worksheet",
                                  }),
                                  beforeSheets.map((beforeSheet) =>
                                    i.jsx(
                                      "option",
                                      {
                                        value: beforeSheet.id,
                                        children: beforeSheet.name,
                                      },
                                      beforeSheet.id,
                                    ),
                                  ),
                                ],
                              }),
                            ],
                          }),
                          i.jsxs("label", {
                            children: [
                              i.jsx(LocalizedText, {
                                children: "Before range ",
                              }),
                              mappingIndex + 1,
                              i.jsx("input", {
                                value: mapping.before.range,
                                onChange: (event) =>
                                  X$1(mappingIndex, {
                                    ...mapping,
                                    before: {
                                      ...mapping.before,
                                      range: event.target.value,
                                    },
                                  }),
                              }),
                            ],
                          }),
                          i.jsxs("label", {
                            children: [
                              i.jsx(LocalizedText, {
                                children: "After worksheet ",
                              }),
                              mappingIndex + 1,
                              i.jsxs("select", {
                                value: mapping.after.sheetId,
                                onChange: (event) =>
                                  X$1(mappingIndex, {
                                    ...mapping,
                                    after: {
                                      ...mapping.after,
                                      sheetId: event.target.value,
                                    },
                                  }),
                                children: [
                                  i.jsx("option", {
                                    value: "",
                                    children: "Choose worksheet",
                                  }),
                                  afterSheets.map((afterSheet) =>
                                    i.jsx(
                                      "option",
                                      {
                                        value: afterSheet.id,
                                        children: afterSheet.name,
                                      },
                                      afterSheet.id,
                                    ),
                                  ),
                                ],
                              }),
                            ],
                          }),
                          i.jsxs("label", {
                            children: [
                              i.jsx(LocalizedText, {
                                children: "After range ",
                              }),
                              mappingIndex + 1,
                              i.jsx("input", {
                                value: mapping.after.range,
                                onChange: (event) =>
                                  X$1(mappingIndex, {
                                    ...mapping,
                                    after: {
                                      ...mapping.after,
                                      range: event.target.value,
                                    },
                                  }),
                              }),
                            ],
                          }),
                          i.jsxs("label", {
                            children: [
                              i.jsx(LocalizedText, {
                                children: "Row matching ",
                              }),
                              mappingIndex + 1,
                              i.jsxs("select", {
                                value: mapping.rowMatch,
                                onChange: (event) =>
                                  X$1(mappingIndex, {
                                    ...mapping,
                                    rowMatch: event.target.value,
                                    keys:
                                      event.target.value === "keys"
                                        ? [{ before: "", after: "" }]
                                        : undefined,
                                  }),
                                children: [
                                  i.jsx("option", {
                                    value: "keys",
                                    children: "Exact typed unique keys",
                                  }),
                                  i.jsx("option", {
                                    value: "position",
                                    children: "Explicit row position",
                                  }),
                                ],
                              }),
                            ],
                          }),
                          i.jsxs("label", {
                            children: [
                              i.jsx("input", {
                                type: "checkbox",
                                checked: mapping.headers,
                                onChange: (event) =>
                                  X$1(mappingIndex, {
                                    ...mapping,
                                    headers: event.target.checked,
                                  }),
                              }),
                              i.jsx(LocalizedText, {
                                children: "First row contains headers ",
                              }),
                              mappingIndex + 1,
                            ],
                          }),
                        ],
                      }),
                      mapping.rowMatch === "keys" &&
                        i.jsxs(i.Fragment, {
                          children: [
                            i.jsx("p", {
                              children: i.jsx(LocalizedText, {
                                children:
                                  "Enter exact header text. Without headers, use relative column labels A, B, C. Values of different types, case or whitespace do not match.",
                              }),
                            }),
                            mapping.keys?.map((key, keyIndex) =>
                              i.jsxs(
                                "div",
                                {
                                  className: "wc-keys",
                                  children: [
                                    i.jsxs("label", {
                                      children: [
                                        i.jsx(LocalizedText, {
                                          children: "Before key ",
                                        }),
                                        mappingIndex + 1,
                                        ".",
                                        keyIndex + 1,
                                        i.jsx("input", {
                                          value: key.before,
                                          onChange: (event) =>
                                            X$1(mappingIndex, {
                                              ...mapping,
                                              keys: mapping.keys.map(
                                                (key, beforeKeyPosition) =>
                                                  beforeKeyPosition === keyIndex
                                                    ? {
                                                        ...key,
                                                        before:
                                                          event.target.value,
                                                      }
                                                    : key,
                                              ),
                                            }),
                                        }),
                                      ],
                                    }),
                                    i.jsxs("label", {
                                      children: [
                                        i.jsx(LocalizedText, {
                                          children: "After key ",
                                        }),
                                        mappingIndex + 1,
                                        ".",
                                        keyIndex + 1,
                                        i.jsx("input", {
                                          value: key.after,
                                          onChange: (event) =>
                                            X$1(mappingIndex, {
                                              ...mapping,
                                              keys: mapping.keys.map(
                                                (key, afterKeyPosition) =>
                                                  afterKeyPosition === keyIndex
                                                    ? {
                                                        ...key,
                                                        after:
                                                          event.target.value,
                                                      }
                                                    : key,
                                              ),
                                            }),
                                        }),
                                      ],
                                    }),
                                    i.jsxs("button", {
                                      disabled: mapping.keys.length < 2,
                                      onClick: () =>
                                        X$1(mappingIndex, {
                                          ...mapping,
                                          keys: mapping.keys.filter(
                                            (key, candidateKeyIndex) =>
                                              candidateKeyIndex !== keyIndex,
                                          ),
                                        }),
                                      children: [
                                        i.jsx(LocalizedText, {
                                          children: "Remove key ",
                                        }),
                                        mappingIndex + 1,
                                        ".",
                                        keyIndex + 1,
                                      ],
                                    }),
                                  ],
                                },
                                keyIndex,
                              ),
                            ),
                            i.jsxs("button", {
                              disabled: (mapping.keys?.length ?? 0) >= 8,
                              onClick: () =>
                                X$1(mappingIndex, {
                                  ...mapping,
                                  keys: [
                                    ...(mapping.keys ?? []),
                                    { before: "", after: "" },
                                  ],
                                }),
                              children: [
                                i.jsx(LocalizedText, {
                                  children: "Add key to mapping ",
                                }),
                                mappingIndex + 1,
                              ],
                            }),
                          ],
                        }),
                      i.jsxs("button", {
                        disabled: mappings.length < 2,
                        onClick: () => {
                          (setMappings(
                            mappings.filter(
                              (mapping, candidateMappingIndex) =>
                                candidateMappingIndex !== mappingIndex,
                            ),
                          ),
                            clearResult());
                        },
                        children: [
                          i.jsx(LocalizedText, { children: "Remove mapping " }),
                          mappingIndex + 1,
                        ],
                      }),
                    ],
                  },
                  mappingIndex,
                ),
              ),
            i.jsxs("div", {
              className: "wc-actions",
              children: [
                i.jsx("button", {
                  disabled: !workbook || mappings.length >= 20,
                  onClick: () => {
                    workbook &&
                      (setMappings([
                        ...mappings,
                        createDefaultMapping(workbook),
                      ]),
                      clearResult());
                  },
                  children: i.jsx(LocalizedText, {
                    children: "Add worksheet mapping",
                  }),
                }),
                i.jsx("button", {
                  disabled: !workbook || isReadingBaseline,
                  onClick: onClick,
                  children: i.jsx(LocalizedText, {
                    children: "Compare selected workbooks",
                  }),
                }),
              ],
            }),
            errorMessage &&
              errorStateKey === inputKey &&
              i.jsx("p", { role: "alert", children: errorMessage }),
            comparisonResult &&
              !isResultCurrent &&
              i.jsx("p", {
                role: "status",
                children: i.jsx(LocalizedText, {
                  children:
                    "Comparison details are withheld because the workbook, selection, settings or access changed. Compare again.",
                }),
              }),
            isResultCurrent &&
              i.jsxs("section", {
                "aria-label": text("Workbook audit comparison report"),
                className: "wc-report",
                lang: language("Workbook audit comparison report"),
                "data-ui-label": "Workbook audit comparison report",
                children: [
                  i.jsx("h3", {
                    children: i.jsx(LocalizedText, {
                      children: "Comparison report",
                    }),
                  }),
                  i.jsxs("p", {
                    children: [
                      comparisonResult.counts.beforeRows,
                      " ",
                      i.jsx(LocalizedText, { children: "before records ·" }),
                      " ",
                      comparisonResult.counts.afterRows,
                      " ",
                      i.jsx(LocalizedText, { children: "after records ·" }),
                      " ",
                      comparisonResult.counts.matchedRows,
                      " ",
                      i.jsx(LocalizedText, { children: "matched ·" }),
                      " ",
                      comparisonResult.counts.insertedRows,
                      " ",
                      i.jsx(LocalizedText, { children: "inserted ·" }),
                      " ",
                      comparisonResult.counts.deletedRows,
                      " ",
                      i.jsx(LocalizedText, { children: "deleted ·" }),
                      " ",
                      comparisonResult.counts.unresolvedBeforeRows,
                      "/",
                      comparisonResult.counts.unresolvedAfterRows,
                      " ",
                      i.jsx(LocalizedText, {
                        children: "unresolved before/after.",
                      }),
                    ],
                  }),
                  i.jsxs("p", {
                    children: [
                      comparisonResult.coverage.selectedCells,
                      " ",
                      i.jsx(LocalizedText, { children: "selected cells." }),
                      " ",
                      comparisonResult.counts.scannedCells,
                      " ",
                      i.jsx(LocalizedText, {
                        children: "permitted stored cells /",
                      }),
                      " ",
                      comparisonResult.counts.scannedFormulas,
                      " ",
                      i.jsx(LocalizedText, {
                        children: "formulas inspected statically.",
                      }),
                      " ",
                      comparisonResult.findings.length,
                      " ",
                      i.jsx(LocalizedText, { children: "findings." }),
                    ],
                  }),
                  i.jsxs("details", {
                    children: [
                      i.jsx("summary", {
                        children: i.jsx(LocalizedText, {
                          children: "Match rules and coverage",
                        }),
                      }),
                      i.jsx("ul", {
                        children: comparisonResult.coverage.notes.map((note) =>
                          i.jsx("li", { children: note }, note),
                        ),
                      }),
                      i.jsxs("p", {
                        children: [
                          i.jsx(LocalizedText, {
                            children: "Before workbook: ",
                          }),
                          comparisonResult.before.id,
                          i.jsx(LocalizedText, { children: ", revision" }),
                          " ",
                          comparisonResult.before.revision,
                          i.jsx(LocalizedText, {
                            children: ". After workbook: ",
                          }),
                          comparisonResult.after.id,
                          i.jsx(LocalizedText, { children: ", revision " }),
                          comparisonResult.after.revision,
                          ".",
                        ],
                      }),
                      comparisonResult.request.mappings.map(
                        (mapping, requestMappingIndex) =>
                          i.jsxs(
                            "p",
                            {
                              children: [
                                i.jsx(LocalizedText, { children: "Mapping " }),
                                requestMappingIndex + 1,
                                ": ",
                                mapping.before.sheetId,
                                "!",
                                mapping.before.range,
                                " →",
                                " ",
                                mapping.after.sheetId,
                                "!",
                                mapping.after.range,
                                "; ",
                                mapping.rowMatch,
                                ";",
                                " ",
                                mapping.headers ? "headers" : "no headers",
                                ";",
                                " ",
                                mapping.keys
                                  ?.map((key) => `${key.before} → ${key.after}`)
                                  .join(", "),
                                ".",
                              ],
                            },
                            requestMappingIndex,
                          ),
                      ),
                    ],
                  }),
                  i.jsxs("label", {
                    children: [
                      i.jsx(LocalizedText, {
                        children: "Comparison finding category",
                      }),
                      i.jsx("select", {
                        value: findingKindFilter,
                        onChange: (event) => {
                          (setFindingKindFilter(event.target.value),
                            setPageIndex(0));
                        },
                        children: comparisonFindingKinds.map(
                          (comparisonFindingKind) =>
                            i.jsxs(
                              "option",
                              {
                                value: comparisonFindingKind,
                                children: [
                                  comparisonFindingKind,
                                  " (",
                                  comparisonFindingKind === "all"
                                    ? comparisonResult.findings.length
                                    : comparisonResult.findings.filter(
                                        (finding) =>
                                          finding.kind ===
                                          comparisonFindingKind,
                                      ).length,
                                  ")",
                                ],
                              },
                              comparisonFindingKind,
                            ),
                        ),
                      }),
                    ],
                  }),
                  i.jsxs("div", {
                    className: "wc-scroll",
                    tabIndex: 0,
                    role: "region",
                    "aria-label": text("Scrollable comparison findings"),
                    lang: language("Scrollable comparison findings"),
                    "data-ui-label": "Scrollable comparison findings",
                    children: [
                      i.jsxs("table", {
                        children: [
                          i.jsx("thead", {
                            children: i.jsxs("tr", {
                              children: [
                                i.jsx("th", {
                                  children: i.jsx(LocalizedText, {
                                    children: "Finding",
                                  }),
                                }),
                                i.jsx("th", {
                                  children: i.jsx(LocalizedText, {
                                    children: "Before address",
                                  }),
                                }),
                                i.jsx("th", {
                                  children: i.jsx(LocalizedText, {
                                    children: "Before raw value",
                                  }),
                                }),
                                i.jsx("th", {
                                  children: i.jsx(LocalizedText, {
                                    children: "After address",
                                  }),
                                }),
                                i.jsx("th", {
                                  children: i.jsx(LocalizedText, {
                                    children: "After raw value",
                                  }),
                                }),
                                i.jsx("th", {
                                  children: i.jsx(LocalizedText, {
                                    children: "Explanation",
                                  }),
                                }),
                              ],
                            }),
                          }),
                          i.jsx("tbody", {
                            children: visibleFindings
                              .slice(currentPage * 25, (currentPage + 1) * 25)
                              .map((pageFinding, findingIndex) =>
                                i.jsxs(
                                  "tr",
                                  {
                                    children: [
                                      i.jsxs("td", {
                                        children: [
                                          pageFinding.kind,
                                          i.jsx("small", {
                                            children: pageFinding.confidence,
                                          }),
                                        ],
                                      }),
                                      i.jsx("td", {
                                        children: pageFinding.before
                                          ? `${pageFinding.before.sheetName}!${pageFinding.before.address}`
                                          : "—",
                                      }),
                                      i.jsx("td", {
                                        children: formatComparisonValue(
                                          pageFinding.beforeValue,
                                        ),
                                      }),
                                      i.jsx("td", {
                                        children: pageFinding.after
                                          ? `${pageFinding.after.sheetName}!${pageFinding.after.address}`
                                          : "—",
                                      }),
                                      i.jsx("td", {
                                        children: formatComparisonValue(
                                          pageFinding.afterValue,
                                        ),
                                      }),
                                      i.jsx("td", {
                                        children: pageFinding.message,
                                      }),
                                    ],
                                  },
                                  `${findingKindFilter}-${currentPage}-${findingIndex}`,
                                ),
                              ),
                          }),
                        ],
                      }),
                      !visibleFindings.length &&
                        i.jsx("p", {
                          children: i.jsx(LocalizedText, {
                            children:
                              "No findings in this category within the declared scope.",
                          }),
                        }),
                    ],
                  }),
                  i.jsxs("div", {
                    className: "wc-actions",
                    children: [
                      i.jsx("button", {
                        disabled: currentPage === 0,
                        onClick: () => setPageIndex(currentPage - 1),
                        children: i.jsx(LocalizedText, {
                          children: "Previous findings",
                        }),
                      }),
                      i.jsxs("span", {
                        children: [
                          i.jsx(LocalizedText, { children: "Page " }),
                          currentPage + 1,
                          " ",
                          i.jsx(LocalizedText, { children: "of " }),
                          pageCount,
                          " · ",
                          visibleFindings.length,
                          " ",
                          i.jsx(LocalizedText, { children: "findings" }),
                        ],
                      }),
                      i.jsx("button", {
                        disabled: currentPage + 1 >= pageCount,
                        onClick: () => setPageIndex(currentPage + 1),
                        children: i.jsx(LocalizedText, {
                          children: "Next findings",
                        }),
                      }),
                    ],
                  }),
                ],
              }),
            i.jsx("p", {
              className: "wc-limit",
              children: i.jsx(LocalizedText, {
                children:
                  "Limits: 20 mappings, 10,000 selected cells, 1,000 rows and 100 columns per range; 20,000 permitted stored cells inspected. Large or incomplete comparisons return an explicit diagnostic.",
              }),
            }),
          ],
        }),
        i.jsx("footer", {
          children: i.jsx("button", {
            onClick: closeDialog,
            children: i.jsx(LocalizedText, { children: "Close comparison" }),
          }),
        }),
      ],
    }),
  });
}
const comparisonLimits = {
  maxQueryLength: 256,
  maxResults: 1e5,
  maxCalculatedCells: 1e4,
  maxTextCharacters: 1e7,
  maxPatternWork: 3e7,
  maxSpecialCells: 500,
  maxSpecialAreas: 256,
  maxBlankScan: 1e5,
};

function normalizeSearchRequest(workbook, r) {
  const sheetOrThrow = findSheetOrThrow(workbook, r.sheetId),
    selectionAreas = parseSelectionAreas(r.ranges ?? [r.range]),
    parsedAreas = selectionAreas.map(parseRangeAddress);
  parseRangeAddress(r.range);
  const scope = r.scope ?? "sheet",
    order = r.order ?? "rows";
  return (
    (!["selection", "sheet", "workbook"].includes(scope) ||
      !["rows", "columns"].includes(order)) &&
      throwWithMessage("Choose a valid search scope and order."),
    r.includeHidden !== undefined &&
      typeof r.includeHidden != "boolean" &&
      throwWithMessage("Include hidden must be boolean."),
    parsedAreas.some(
      (area) =>
        area.lastRow >= sheetOrThrow.rowCount ||
        area.lastCol >= sheetOrThrow.colCount,
    ) && throwWithMessage("The search selection exceeds the worksheet."),
    {
      ...r,
      ranges: selectionAreas,
      scope: scope,
      order: order,
      includeHidden: r.includeHidden ?? false,
    }
  );
}
function normalizeSearchQuery(workbook, request) {
  const searchRequest = normalizeSearchRequest(workbook, request),
    lookIn = request.lookIn ?? "raw";
  ((typeof request.query != "string" ||
    !request.query.length ||
    request.query.length > comparisonLimits.maxQueryLength) &&
    throwWithMessage(
      "Enter 1–256 characters to find. Use Go To Special to select blank cells.",
    ),
    ["raw", "formulas", "calculated", "formatted"].includes(lookIn) ||
      throwWithMessage("Choose a supported Look in option."));
  for (const switchValue of [
    request.matchCase,
    request.wholeCell,
    request.wildcards,
  ])
    switchValue !== undefined &&
      typeof switchValue != "boolean" &&
      throwWithMessage("Search switches must be boolean.");
  return {
    ...searchRequest,
    query: request.query,
    lookIn: lookIn,
    matchCase: request.matchCase ?? false,
    wholeCell: request.wholeCell ?? false,
    wildcards: request.wildcards ?? true,
  };
}
const createSearchBudget = () => ({
  text: 0,
  work: 0,
  calculated: 0,
  scanned: 0,
});
function compileSearchPattern(searchQuery) {
  const characters = Array.from(searchQuery.query),
    tokens = [];
  for (let charIndex = 0; charIndex < characters.length; charIndex++) {
    const character = characters[charIndex];
    searchQuery.wildcards &&
    character === "~" &&
    ["~", "?", "*"].includes(characters[charIndex + 1])
      ? tokens.push({ kind: "literal", value: characters[++charIndex] })
      : searchQuery.wildcards && character === "*"
        ? tokens.at(-1)?.kind !== "star" && tokens.push({ kind: "star" })
        : searchQuery.wildcards && character === "?"
          ? tokens.push({ kind: "any" })
          : tokens.push({ kind: "literal", value: character });
  }
  return tokens;
}
function matchTextWithBudget(text, options, tokens, searchBudget) {
  const characters = Array.from(text),
    characterCount = characters.length;
  if (
    ((searchBudget.text += text.length),
    (searchBudget.work += (characterCount + 1) * tokens.length),
    (searchBudget.text > comparisonLimits.maxTextCharacters ||
      searchBudget.work > comparisonLimits.maxPatternWork) &&
      throwWithMessage(
        "This search exceeds the text-comparison budget. Narrow the selection or simplify the search.",
      ),
    !characterCount)
  )
    return [];
  let matchEnds = Int32Array.from(
    { length: characterCount + 1 },
    (l, index) => index,
  );
  for (let tokenIndex = tokens.length - 1; tokenIndex >= 0; tokenIndex--) {
    const token = tokens[tokenIndex],
      nextMatchEnds = new Int32Array(characterCount + 1).fill(-1);
    if (token.kind === "star") {
      nextMatchEnds[characterCount] = matchEnds[characterCount];
      for (
        let starCharPosition = characterCount - 1;
        starCharPosition >= 0;
        starCharPosition--
      )
        nextMatchEnds[starCharPosition] = Math.max(
          matchEnds[starCharPosition],
          nextMatchEnds[starCharPosition + 1],
        );
    } else
      for (
        let charPosition = characterCount - 1;
        charPosition >= 0;
        charPosition--
      )
        (token.kind === "any" ||
          (options.matchCase
            ? characters[charPosition] === token.value
            : characters[charPosition].toLowerCase() ===
              token.value.toLowerCase())) &&
          (nextMatchEnds[charPosition] = matchEnds[charPosition + 1]);
    matchEnds = nextMatchEnds;
  }
  const codeUnitOffsets = [0];
  for (const character of characters)
    codeUnitOffsets.push(codeUnitOffsets.at(-1) + character.length);
  if (options.wholeCell)
    return matchEnds[0] === characterCount ? [[0, text.length]] : [];
  const matchRanges = [];
  for (let charIndex = 0; charIndex < characterCount;)
    matchEnds[charIndex] > charIndex
      ? (matchRanges.push([
          codeUnitOffsets[charIndex],
          codeUnitOffsets[matchEnds[charIndex]],
        ]),
        (charIndex = matchEnds[charIndex]))
      : charIndex++;
  return matchRanges;
}
function describeHiddenState(sheetOrThrow) {
  const axisVisibilityIndex = createAxisVisibilityIndex(
      sheetOrThrow.rowCount,
      cn(sheetOrThrow, "row"),
    ),
    columnVisibilityIndex = createAxisVisibilityIndex(
      sheetOrThrow.colCount,
      cn(sheetOrThrow, "column"),
    );
  return (rowIndex, columnIndex) =>
    sheetOrThrow.hidden
      ? sheetOrThrow.veryHidden
        ? "very hidden sheet"
        : "hidden sheet"
      : axisVisibilityIndex.isHidden(rowIndex)
        ? "hidden row"
        : columnVisibilityIndex.isHidden(columnIndex)
          ? "hidden column"
          : undefined;
}
const rangeContainsCell = (row, column, range) =>
  row >= range.firstRow &&
  row <= range.lastRow &&
  column >= range.firstCol &&
  column <= range.lastCol;
function compareCellPositions(order, leftCell, rightCell) {
  return order === "columns"
    ? leftCell.col - rightCell.col || leftCell.row - rightCell.row
    : leftCell.row - rightCell.row || leftCell.col - rightCell.col;
}
function searchWorkbookCells(workbook, searchRequest) {
  const searchQuery = normalizeSearchQuery(workbook, searchRequest),
    searchBudget = createSearchBudget(),
    searchPattern = compileSearchPattern(searchQuery),
    results = [],
    workbookWithSettings = calculateWorkbookWithSettings(workbook),
    sheetsToSearch =
      searchQuery.scope === "workbook"
        ? workbook.sheets
        : [findSheetOrThrow(workbook, searchQuery.sheetId)],
    parsedRanges = searchQuery.ranges.map(parseRangeAddress),
    shouldUseCalculatedValues =
      searchQuery.lookIn === "calculated" || searchQuery.lookIn === "formatted";
  for (const sheet of sheetsToSearch) {
    if (sheet.hidden && !searchQuery.includeHidden) continue;
    const candidateAddresses = new Set(Object.keys(sheet.cells)),
      hiddenReasonFor = describeHiddenState(sheet),
      sheetMatches = [];
    if (shouldUseCalculatedValues)
      for (const spillEntry of workbookWithSettings.getSpillEntries(sheet.id))
        candidateAddresses.add(spillEntry.address);
    for (const address of candidateAddresses) {
      ++searchBudget.scanned > workbookLimits.maxCells &&
        throwWithMessage(
          "This search exceeds 100,000 populated cells. Narrow its scope.",
        );
      const { row, col } = _t(address),
        hiddenReason = hiddenReasonFor(row, col),
        cell = sheet.cells[address];
      if (
        (searchQuery.scope === "selection" &&
          !parsedRanges.some((range) => rangeContainsCell(row, col, range))) ||
        (hiddenReason && !searchQuery.includeHidden)
      )
        continue;
      const isFormula = isFormulaCellValue(cell);
      if (searchQuery.lookIn === "formulas" && !isFormula) continue;
      let cellValue = cell?.raw ?? null;
      shouldUseCalculatedValues &&
        (++searchBudget.calculated > comparisonLimits.maxCalculatedCells &&
          throwWithMessage(
            "Calculated searches support at most 10,000 cells. Narrow the selection or use Raw contents.",
          ),
        (cellValue = workbookWithSettings(sheet.id, address)));
      const cellText =
          searchQuery.lookIn === "formatted"
            ? formatNumberWithPattern(
                cellValue,
                cell,
                workbook.datePolicy?.dateSystem,
              )
            : coerceCellValueToText(cellValue),
        textMatches = matchTextWithBudget(
          cellText,
          searchQuery,
          searchPattern,
          searchBudget,
        );
      textMatches.length &&
        (sheetMatches.push({
          sheetId: sheet.id,
          sheetName: sheet.name,
          address: address,
          row,
          col,
          text: cellText,
          raw: cell?.raw ?? null,
          isFormula: isFormula,
          protected: !!sheet.protection?.enabled,
          hidden: !!hiddenReason,
          ...(hiddenReason ? { hiddenReason: hiddenReason } : {}),
          occurrences: textMatches.length,
        }),
        results.length + sheetMatches.length > comparisonLimits.maxResults &&
          throwWithMessage("Too many matching cells. Narrow the search."));
    }
    (sheetMatches.sort((leftResult, rightResult) =>
      compareCellPositions(searchQuery.order, leftResult, rightResult),
    ),
      results.push(...sheetMatches));
  }
  return eI({
    workbookId: workbook.id,
    revision: workbook.revision,
    snapshotGuard: JSON.stringify(workbook),
    sheetOrder: workbook.sheets.map((sheet) => sheet.id),
    options: { ...searchQuery },
    total: results.length,
    scannedCells: searchBudget.scanned,
    results: results,
  });
}
function normalizeSearchPage(searchResult, pageIndex = 0, pageSize = 25) {
  return (
    (!Number.isSafeInteger(pageIndex) ||
      pageIndex < 0 ||
      !Number.isSafeInteger(pageSize) ||
      pageSize < 1 ||
      pageSize > 200) &&
      throwWithMessage(
        "Search pages require a nonnegative page and 1–200 results per page.",
      ),
    searchResult.results.slice(pageIndex * pageSize, (pageIndex + 1) * pageSize)
  );
}
function findAdjacentSearchResult(
  searchResult,
  currentResult,
  direction = "next",
) {
  if (
    (["next", "previous"].includes(direction) ||
      throwWithMessage("Choose next or previous."),
    !searchResult.total)
  )
    return null;
  const currentIndex = currentResult
    ? searchResult.results.findIndex(
        (result) =>
          result.sheetId === currentResult.sheetId &&
          result.address === currentResult.address.toUpperCase(),
      )
    : -1;
  if (currentIndex >= 0)
    return searchResult.results[
      (currentIndex + (direction === "next" ? 1 : searchResult.total - 1)) %
        searchResult.total
    ];
  if (currentResult) {
    const currentPosition = _t(currentResult.address),
      sheetIndexById = new Map(
        searchResult.sheetOrder.map((sheetId, sheetIndex) => [
          sheetId,
          sheetIndex,
        ]),
      ),
      currentSheetIndex = sheetIndexById.get(currentResult.sheetId);
    if (currentSheetIndex !== undefined) {
      const compareWithCurrent = (candidate) =>
        sheetIndexById.get(candidate.sheetId) - currentSheetIndex ||
        compareCellPositions(
          searchResult.options.order,
          candidate,
          currentPosition,
        );
      return direction === "next"
        ? (searchResult.results.find(
            (result) => compareWithCurrent(result) > 0,
          ) ?? searchResult.results[0])
        : ([...searchResult.results]
            .reverse()
            .find(
              (precedingCandidate) =>
                compareWithCurrent(precedingCandidate) < 0,
            ) ?? searchResult.results.at(-1));
    }
  }
  return direction === "next"
    ? searchResult.results[0]
    : searchResult.results.at(-1);
}
function toComparableCellValue(cell, enteredValue) {
  let raw = enteredValue;
  enteredValue.length
    ? typeof cell.raw == "number" &&
      /^[+-]?(?:\d+\.?\d*|\.\d+)(?:e[+-]?\d+)?$/i.test(enteredValue) &&
      Number.isFinite(Number(enteredValue))
      ? (raw = Number(enteredValue))
      : typeof cell.raw == "boolean" &&
        /^(true|false)$/i.test(enteredValue) &&
        (raw = enteredValue.toLowerCase() === "true")
    : (raw = null);
  const comparableCell = { ...cell, raw: raw };
  return (
    typeof raw == "string" && raw.startsWith("=") && !isFormulaCellValue(cell)
      ? (comparableCell.literal = true)
      : (typeof raw == "string" && raw.startsWith("=")) ||
        delete comparableCell.literal,
    normalizeCellRecordInput(comparableCell)
  );
}

export {
  compareCellPositions,
  compileSearchPattern,
  searchWorkbookCells,
  normalizeSearchPage,
  findAdjacentSearchResult,
  normalizeSearchRequest,
  describeHiddenState,
  createSearchBudget,
  toComparableCellValue,
  comparisonLimits,
  matchTextWithBudget,
  WorkbookComparisonDialog,
  rangeContainsCell,
};
