import {
  isNullOrEmptyValue,
  serializeTypedKeyValues,
  failWithMessage,
} from "./03-workbook-model.js";
import { formatCellAddress } from "./04-cell-input-paste-and-fill.js";
import {
  transformationRecipeLimits,
  fingerprintCanonicalValue,
  normalizeDatePolicy,
  dateSystemMaxSerial,
  parseDateText,
} from "./01-schemas-date-functions-and-validation.js";
import { workbookLimits } from "./02-text-and-financial-functions.js";
import {
  textJoinSpecSchema,
  validateFieldTransformRecipe,
  createWorksheetRecord,
  findRowOrigin,
  calculateWorkbookWithSettings,
  validateFieldTransformEvidence,
  textInferenceLimits,
  sourceRowValueSchema,
  _t,
  parseRangeAddress,
  normalizeCellRecordInput,
  normalizeWorkbookInput,
  columnIndexToLetters,
  checkDataValidationBatch,
} from "./03-workbook-model.js";
import {
  readSourceHeaders,
  matchKeyPopulations,
  inspectFormulaDependencies,
} from "./12-consolidation-and-dependency-inspection.js";
import {
  validateTextToColumnsRequest,
  assertSelectionInAuthorizedSheet,
  splitDelimitedCell,
  parseImportedCellValue,
  rangesOverlap,
  dataModelBudgets,
  formatCellTextValue,
  validateTextToColumnsPermissions,
  assertOutputNameUnreferenced,
} from "./15-data-model-dialog.js";

function discoverArrayFormulaAnchors(workbook, sheetRanges, context) {
  const arrayAnchors = [];
  let formulaCellCount = 0,
    formulaCharacterCount = 0;
  for (const sheet of new Map(
    sheetRanges.map((sheetRange) => [sheetRange.sheet.id, sheetRange.sheet]),
  ).values())
    for (const [anchor, cell] of Object.entries(sheet.cells)) {
      if (
        typeof cell.raw != "string" ||
        !cell.raw.startsWith("=") ||
        cell.literal
      )
        continue;
      (++formulaCellCount > dataModelBudgets.formulaCells ||
        (formulaCharacterCount += cell.raw.length) >
          dataModelBudgets.formulaCharacters) &&
        failWithMessage(
          "Array checking supports at most 1,000 formulas and 100,000 formula characters on the participating worksheets.",
        );
      const anchorPosition = _t(anchor);
      sheetRanges.some((rangeEntry) => {
        if (rangeEntry.sheet.id !== sheet.id) return false;
        const rangeAddress = parseRangeAddress(rangeEntry.range),
          overlapFirstRow = Math.max(anchorPosition.row, rangeAddress.firstRow),
          overlapFirstCol = Math.max(anchorPosition.col, rangeAddress.firstCol);
        return (
          overlapFirstRow <= rangeAddress.lastRow &&
          overlapFirstCol <= rangeAddress.lastCol &&
          (overlapFirstRow - anchorPosition.row + 1) *
            (overlapFirstCol - anchorPosition.col + 1) <=
            1e4
        );
      }) && arrayAnchors.push({ sheet: sheet, anchor: anchor });
    }
  if (
    (arrayAnchors.length > dataModelBudgets.discoveryAnchors &&
      failWithMessage(
        "The selection has more than 50 possible array anchors. Choose a new output sheet or a smaller reviewed region.",
      ),
    !arrayAnchors.length)
  )
    return;
  const dependencyKeys = new Set(),
    dependencyKey = (keySheetId, keyAddress) =>
      JSON.stringify([keySheetId, keyAddress]),
    dependencyWorkbook = {
      ...workbook,
      sheets: workbook.sheets.map((sheet) =>
        sheet.veryHidden ? { ...sheet, hidden: true } : sheet,
      ),
    };
  for (const arrayAnchor of arrayAnchors) {
    assertSelectionInAuthorizedSheet(
      arrayAnchor.sheet,
      arrayAnchor.anchor,
      context,
    );
    const dependencyReport = inspectFormulaDependencies(dependencyWorkbook, {
        sheetId: arrayAnchor.sheet.id,
        address: arrayAnchor.anchor,
        direction: "precedents",
        depth: 8,
        includeHidden: context.includeHidden,
        allowedSheetIds: context.allowedSheetIds,
        limits: { maxFormulaCells: 1e3, maxNodes: 500, maxReferences: 2e3 },
      }),
      closureNodes = [dependencyReport.target, ...dependencyReport.nodes],
      closureKeys = new Set(closureNodes.map((event) => event.key));
    (dependencyReport.coverage.truncated ||
      dependencyReport.coverage.depthLimited ||
      dependencyReport.cycles.length ||
      dependencyReport.issues.some((issue) => closureKeys.has(issue.from))) &&
      failWithMessage(
        "Array discovery requires a complete, authorized static dependency closure. Resolve hidden, dynamic, cyclic or unsupported formula dependencies before splitting.",
      );
    for (const closureNode of closureNodes) {
      const sheet = workbook.sheets.find(
        (sheet) => sheet.id === closureNode.sheetId,
      );
      (assertSelectionInAuthorizedSheet(sheet, closureNode.address, context),
        (!sheet.cells[closureNode.address] ||
          sheet.cells[closureNode.address].raw === null) &&
          failWithMessage(
            "Array discovery cannot infer blank or possible spill-child dependencies. Use a reviewed values snapshot before splitting.",
          ),
        dependencyKeys.add(
          dependencyKey(closureNode.sheetId, closureNode.address),
        ),
        dependencyKeys.size > dataModelBudgets.dependencyCells &&
          failWithMessage(
            "Array discovery exceeds 2,000 dependency cells. Use a smaller reviewed values snapshot.",
          ));
    }
  }
  const allowedWorkbook = {
      ...workbook,
      sheets: workbook.sheets
        .filter((sheet) => context.allowedSheetIds.includes(sheet.id))
        .map((allowedSheet) => ({
          ...allowedSheet,
          cells: Object.fromEntries(
            Object.entries(allowedSheet.cells).filter(([cellAddress]) =>
              dependencyKeys.has(dependencyKey(allowedSheet.id, cellAddress)),
            ),
          ),
        })),
    },
    workbookWithSettings = calculateWorkbookWithSettings(allowedWorkbook);
  for (const { sheet, anchor } of arrayAnchors) {
    const spillRange = workbookWithSettings.getSpillRange(sheet.id, anchor);
    spillRange &&
      sheetRanges.some(
        (selectionEntry) =>
          selectionEntry.sheet.id === sheet.id &&
          rangesOverlap(selectionEntry.range, spillRange),
      ) &&
      failWithMessage(
        "The selected source or destination intersects a generated array. Convert its results to ordinary values explicitly before splitting.",
      );
  }
  workbookWithSettings.getOperationalErrors().length &&
    failWithMessage(
      "Array discovery encountered an operational calculation error. Resolve the uncertain formula state before splitting text.",
    );
}
function buildTextSplitResult(workbook, request, permissionOptions = {}) {
  const context = validateTextToColumnsPermissions(workbook, permissionOptions),
    splitRequest = validateTextToColumnsRequest(request),
    sheet = workbook.sheets.find(
      (sheet) => sheet.id === splitRequest.sourceSheetId,
    );
  (sheet || failWithMessage("Choose an existing source worksheet."),
    assertSelectionInAuthorizedSheet(sheet, splitRequest.sourceRange, context));
  const rangeAddress = parseRangeAddress(splitRequest.sourceRange),
    height = rangeAddress.lastRow - rangeAddress.firstRow + 1;
  ((rangeAddress.firstCol !== rangeAddress.lastCol ||
    height > dataModelBudgets.sourceCells) &&
    failWithMessage("Select one source column containing at most 500 cells."),
    sheet.merges?.some((merge) =>
      rangesOverlap(merge, splitRequest.sourceRange),
    ) && failWithMessage("Unmerge the source before splitting text."));
  const sources = [],
    fieldsByRow = [],
    diagnostics = [];
  let maxFieldCount = 0,
    sourceCharacterCount = 0;
  for (
    let firstRow = rangeAddress.firstRow;
    firstRow <= rangeAddress.lastRow;
    firstRow++
  ) {
    const cellAddress = formatCellAddress(firstRow, rangeAddress.firstCol),
      cell = sheet.cells[cellAddress];
    cell &&
      typeof cell.raw == "string" &&
      cell.raw.startsWith("=") &&
      !cell.literal &&
      failWithMessage(
        `Source ${cellAddress} contains a formula. Copy its results to ordinary values explicitly before splitting.`,
      );
    const cellTextValue = formatCellTextValue(cell?.raw ?? null);
    (sourceCharacterCount += cellTextValue.length) >
      dataModelBudgets.sourceCharacters &&
      failWithMessage(
        "Selected source text exceeds the 200,000-character processing budget.",
      );
    let fields;
    try {
      fields = splitDelimitedCell(cellTextValue, splitRequest.delimiter);
    } catch (error) {
      failWithMessage(`Source ${cellAddress}: ${error.message}`);
    }
    ((maxFieldCount = Math.max(maxFieldCount, fields.length)),
      fieldsByRow.push(fields),
      sources.push({
        address: cellAddress,
        ...(cell ? { cell: structuredClone(cell) } : {}),
        text: cellTextValue,
        fields: fields.length,
      }));
  }
  (maxFieldCount * height > dataModelBudgets.outputCells &&
    failWithMessage(
      "Split output may contain at most 500 cells. Select fewer source rows or fields.",
    ),
    (splitRequest.formats?.length ?? 0) > maxFieldCount &&
      failWithMessage(
        "A column conversion refers to a field outside the split result.",
      ));
  const formats = Array.from(
    { length: maxFieldCount },
    (O, fieldPosition) => splitRequest.formats?.[fieldPosition] ?? "text",
  );
  let targetSheet,
    targetSheetName,
    targetOrigin = { row: 0, col: 0 };
  if (splitRequest.destination.kind === "new-sheet")
    ((!context.allowNewSheet || workbook.structureProtection?.enabled) &&
      failWithMessage(
        "Creating an output worksheet is not permitted while workbook structure is protected or the host disallows new worksheets.",
      ),
      (targetSheetName = splitRequest.destination.name),
      workbook.sheets.some(
        (sheet) => sheet.name.toLowerCase() === targetSheetName.toLowerCase(),
      ) &&
        failWithMessage(
          "The output worksheet name is already used. Choose another name.",
        ),
      assertOutputNameUnreferenced(workbook, targetSheetName));
  else {
    const destination = splitRequest.destination;
    ((targetSheet = workbook.sheets.find(
      (sheet) => sheet.id === destination.sheetId,
    )),
      targetSheet || failWithMessage("Choose an existing output worksheet."),
      (targetOrigin = _t(splitRequest.destination.address)),
      (targetSheetName = targetSheet.name),
      targetSheet.protection?.enabled &&
        failWithMessage("The output worksheet is protected."));
  }
  (targetOrigin.row + height > workbookLimits.maxRows ||
    targetOrigin.col + maxFieldCount > workbookLimits.maxCols) &&
    failWithMessage("Split output exceeds the worksheet boundary.");
  const targetRange = `${formatCellAddress(targetOrigin.row, targetOrigin.col)}:${formatCellAddress(targetOrigin.row + height - 1, targetOrigin.col + maxFieldCount - 1)}`;
  (targetSheet &&
    (assertSelectionInAuthorizedSheet(targetSheet, targetRange, context),
    targetSheet.id === sheet.id &&
      rangesOverlap(splitRequest.sourceRange, targetRange) &&
      failWithMessage(
        "Choose a separate destination so the original source cells remain intact.",
      ),
    (targetSheet.tables?.some((table) =>
      rangesOverlap(table.range, targetRange),
    ) ||
      targetSheet.merges?.some((merge) => rangesOverlap(merge, targetRange)) ||
      [...(targetSheet.pivots ?? []), ...(targetSheet.queries ?? [])].some(
        (analysisResult) =>
          rangesOverlap(analysisResult.outputRange, targetRange),
      )) &&
      failWithMessage(
        "Choose an output outside tables, merged areas and generated analysis results.",
      ),
    Object.keys(targetSheet.cells).some((cell) =>
      rangesOverlap(cell, targetRange),
    ) &&
      failWithMessage(
        "The destination contains cells or their formatting. Choose an empty range or new worksheet.",
      ),
    [
      ...(workbook.reviewThreads ?? []),
      ...(workbook.formulaWatches ?? []),
    ].some(
      (anchoredItem) =>
        !anchoredItem.orphaned &&
        anchoredItem.anchor.sheetId === targetSheet.id &&
        rangesOverlap(anchoredItem.anchor.address, targetRange),
    ) &&
      failWithMessage(
        "The destination contains a saved watch or discussion anchor. Choose another range to preserve its evidence context.",
      )),
    discoverArrayFormulaAnchors(
      workbook,
      [
        { sheet: sheet, range: splitRequest.sourceRange },
        ...(targetSheet ? [{ sheet: targetSheet, range: targetRange }] : []),
      ],
      context,
    ));
  const outputs = [];
  for (let recordIndex = 0; recordIndex < height; recordIndex++) {
    fieldsByRow[recordIndex].length !== maxFieldCount &&
      diagnostics.push({
        code: "padded-row",
        sourceAddress: sources[recordIndex].address,
        message: `This record has ${fieldsByRow[recordIndex].length} fields; ${maxFieldCount - fieldsByRow[recordIndex].length} missing fields are padded with empty text.`,
      });
    for (let fieldIndex = 0; fieldIndex < maxFieldCount; fieldIndex++) {
      const field = fieldsByRow[recordIndex][fieldIndex] ?? "",
        format =
          splitRequest.header === "first-row" && recordIndex === 0
            ? "text"
            : formats[fieldIndex],
        importedCellValue = parseImportedCellValue(field, format);
      importedCellValue.precision &&
        diagnostics.push({
          code: "numeric-text-retained",
          sourceAddress: sources[recordIndex].address,
          message: `Field ${fieldIndex + 1} remains text because numeric conversion would exceed the supported finite precision.`,
        });
      const note = `Text to Columns derivation
Workbook ID: ${JSON.stringify(workbook.id)}
Source revision: ${workbook.revision}
Worksheet ID: ${JSON.stringify(sheet.id)}
Worksheet name: ${JSON.stringify(sheet.name)}
Source cell: ${sources[recordIndex].address}
Field: ${fieldIndex + 1}
Conversion: ${format}
${fieldIndex >= fieldsByRow[recordIndex].length ? "Padded missing field." : "Original source cell and evidence remain in place."}`,
        cellRecordInput = normalizeCellRecordInput({
          raw: importedCellValue.raw,
          note: note,
          ...(typeof importedCellValue.raw == "string"
            ? {
                numFmt: "@",
                ...(importedCellValue.raw.startsWith("=")
                  ? { literal: true }
                  : {}),
              }
            : {}),
          ...(splitRequest.header === "first-row" && recordIndex === 0
            ? { style: { bold: true } }
            : {}),
        });
      outputs.push({
        sourceAddress: sources[recordIndex].address,
        address: formatCellAddress(
          targetOrigin.row + recordIndex,
          targetOrigin.col + fieldIndex,
        ),
        field: field,
        padded: fieldIndex >= fieldsByRow[recordIndex].length,
        format: format,
        type:
          typeof importedCellValue.raw == "number"
            ? "number"
            : typeof importedCellValue.raw == "boolean"
              ? "boolean"
              : "text",
        cell: cellRecordInput,
      });
    }
  }
  if (
    (workbook.sheets.reduce(
      (total, existingSheet) => total + Object.keys(existingSheet.cells).length,
      0,
    ) +
      outputs.length >
      workbookLimits.maxCells &&
      failWithMessage(
        "The split would exceed the workbook's 100,000 stored-cell budget.",
      ),
    targetSheet)
  ) {
    const book = structuredClone(workbook),
      sheet = book.sheets.find((sheet) => sheet.id === targetSheet.id);
    for (const output of outputs) sheet.cells[output.address] = output.cell;
    checkDataValidationBatch(
      book,
      outputs.map((output) => ({
        sheetId: sheet.id,
        address: output.address,
        after: output.cell,
      })),
    );
  }
  return {
    snapshot: JSON.stringify(workbook),
    request: splitRequest,
    context: context,
    sourceSheetName: sheet.name,
    targetSheetName: targetSheetName,
    targetRange: targetRange,
    width: maxFieldCount,
    height: height,
    formats: formats,
    sources: sources,
    outputs: outputs,
    diagnostics: diagnostics,
  };
}
function applyTextSplitPreview(workbook, preview, permissionOptions = {}) {
  preview.snapshot !== JSON.stringify(workbook) &&
    failWithMessage(
      "The workbook changed after the split preview. Preview again.",
    );
  const textSplitResult = buildTextSplitResult(
    workbook,
    preview.request,
    permissionOptions,
  );
  JSON.stringify(textSplitResult) !== JSON.stringify(preview) &&
    failWithMessage(
      "The split preview or authorized context changed. Preview again.",
    );
  const nextWorkbook = structuredClone(workbook),
    destination = preview.request.destination,
    targetSheet =
      destination.kind === "new-sheet"
        ? createWorksheetRecord(destination.name)
        : nextWorkbook.sheets.find((sheet) => sheet.id === destination.sheetId);
  destination.kind === "new-sheet" && nextWorkbook.sheets.push(targetSheet);
  for (const output of preview.outputs)
    targetSheet.cells[output.address] = structuredClone(output.cell);
  return (normalizeWorkbookInput(nextWorkbook), nextWorkbook);
}
const coerceCellValueToText = (value) =>
  value === null
    ? ""
    : typeof value == "boolean"
      ? value
        ? "TRUE"
        : "FALSE"
      : String(value);
function applyTextCasing(text, casing) {
  return casing === "upper"
    ? text.toLocaleUpperCase("en-US")
    : casing === "lower"
      ? text.toLocaleLowerCase("en-US")
      : casing === "proper"
        ? text
            .toLocaleLowerCase("en-US")
            .replace(new RegExp("\\p{L}[\\p{L}\\p{M}]*", "gu"), (word) => {
              const [firstCharacter, ...restCharacters] = [...word];
              return (
                firstCharacter.toLocaleUpperCase("en-US") +
                restCharacters.join("")
              );
            })
        : text;
}
function extractInferredFieldPart(rowValues, part) {
  if (part.column >= rowValues.length)
    throw Error("The inferred rule refers to a missing field.");
  const fieldText = coerceCellValueToText(rowValues[part.column]).trim();
  if (part.mode === "whole") return fieldText;
  if (part.mode === "slice") {
    const characters = [...fieldText];
    if (characters.length < part.start + part.length)
      throw Error(
        "The value is too short for the inferred character extraction.",
      );
    return characters.slice(part.start, part.start + part.length).join("");
  }
  const delimiterIndex = fieldText.indexOf(part.delimiter);
  if (part.mode === "before" || part.mode === "after") {
    if (delimiterIndex < 0) throw Error("The inferred separator is absent.");
    return (
      part.mode === "before"
        ? fieldText.slice(0, delimiterIndex)
        : fieldText.slice(delimiterIndex + part.delimiter.length)
    ).trim();
  }
  const tokens =
      part.delimiter === " "
        ? fieldText.split(/ +/)
        : fieldText.split(part.delimiter),
    tokenIndex = part.index < 0 ? tokens.length + part.index : part.index;
  if (tokenIndex < 0 || tokenIndex >= tokens.length || !tokens[tokenIndex])
    throw Error("The inferred token is absent.");
  return tokens[tokenIndex].trim();
}
function applyInferredTextPattern(rowValues, pattern) {
  const patternSpec = textJoinSpecSchema.parse(pattern),
    patternText =
      patternSpec.prefix +
      applyTextCasing(
        patternSpec.parts
          .map((part) => extractInferredFieldPart(rowValues, part))
          .join(patternSpec.separator),
        patternSpec.casing,
      ) +
      patternSpec.suffix;
  if (patternText.length > 32767)
    throw Error("The inferred result exceeds the cell text limit.");
  return patternText;
}
function describeInferredTextRule(pattern) {
  const partsDescription = pattern.parts
    .map(
      (part) =>
        `field ${part.column + 1}${part.mode === "whole" ? "" : part.mode === "token" ? ` token ${part.index < 0 ? "last" : part.index + 1} by ${JSON.stringify(part.delimiter)}` : part.mode === "slice" ? ` characters ${part.start + 1}–${part.start + part.length}` : ` ${part.mode} ${JSON.stringify(part.delimiter)}`}`,
    )
    .join(` + ${JSON.stringify(pattern.separator)} + `);
  return `${pattern.prefix ? JSON.stringify(pattern.prefix) + " + " : ""}${pattern.casing === "none" ? "" : pattern.casing + "("}${partsDescription}${pattern.casing === "none" ? "" : ")"}${pattern.suffix ? " + " + JSON.stringify(pattern.suffix) : ""}`;
}
function inferTextPatternFromExamples(dataRows, fieldColumns, exampleInputs) {
  if (
    !Array.isArray(exampleInputs) ||
    exampleInputs.length < 2 ||
    exampleInputs.length > 10
  )
    throw Error("Provide 2–10 examples from distinct source rows.");
  const examples = exampleInputs.map((rawExample) =>
    sourceRowValueSchema.parse(rawExample),
  );
  if (
    new Set(examples.map((parsedExample) => parsedExample.row)).size !==
    examples.length
  )
    throw Error("Provide examples from distinct source rows.");
  const exampleRows = examples.map((example) => ({
    example: example,
    row: dataRows.find((dataRow) => dataRow.row === example.row),
  }));
  if (exampleRows.some((exampleRowEntry) => !exampleRowEntry.row))
    throw Error("An example row is outside the selected data rows.");
  if (
    !fieldColumns.length ||
    fieldColumns.length > 8 ||
    new Set(fieldColumns).size !== fieldColumns.length ||
    fieldColumns.some(
      (candidateField) =>
        !Number.isInteger(candidateField) ||
        candidateField < 0 ||
        exampleRows.some(
          (exampleEntry) => candidateField >= exampleEntry.row.values.length,
        ),
    )
  )
    throw Error("Choose 1–8 available fields for inference.");
  const singleParts = [],
    makeColumnPart = (column) => ({
      column: column,
      mode: "whole",
      delimiter: " ",
      index: 0,
      start: 0,
      length: 1,
    });
  for (const fieldColumn of fieldColumns) {
    singleParts.push(makeColumnPart(fieldColumn));
    for (const delimiter of [" ", ",", "-", "/", "_", "@"]) {
      for (const tokenIndex of [0, 1, -1])
        singleParts.push({
          ...makeColumnPart(fieldColumn),
          mode: "token",
          delimiter: delimiter,
          index: tokenIndex,
        });
      for (const mode of ["before", "after"])
        singleParts.push({
          ...makeColumnPart(fieldColumn),
          mode: mode,
          delimiter: delimiter,
        });
    }
    for (const sliceLength of [1, 2, 3, 4, 5, 6, 8, 10])
      singleParts.push({
        ...makeColumnPart(fieldColumn),
        mode: "slice",
        length: sliceLength,
      });
  }
  const partPatterns = singleParts.map((singlePart) => ({
    parts: [singlePart],
    separator: "",
  }));
  for (const leftColumn of fieldColumns)
    for (const rightColumn of fieldColumns)
      if (leftColumn !== rightColumn)
        for (const separator of [" ", "", ", ", ".", "-", "_", "/"])
          partPatterns.push({
            parts: [makeColumnPart(leftColumn), makeColumnPart(rightColumn)],
            separator: separator,
          });
  const candidates = [],
    seenPatterns = new Set();
  let searched = 0,
    truncated = false;
  e: for (const partPattern of partPatterns)
    for (const casing of ["none", "upper", "lower", "proper"]) {
      if (++searched > textInferenceLimits.inferencePatterns) {
        (searched--, (truncated = true));
        break e;
      }
      const basePattern = {
        ...partPattern,
        casing: casing,
        prefix: "",
        suffix: "",
      };
      let patternValues;
      try {
        patternValues = exampleRows.map((exampleItem) =>
          applyInferredTextPattern(exampleItem.row.values, basePattern),
        );
      } catch {
        continue;
      }
      if (
        patternValues.some((patternValue) => !patternValue) ||
        new Set(patternValues).size < 2
      )
        continue;
      const value = exampleRows[0].example.value,
        firstValue = patternValues[0];
      let matchStart = value.indexOf(firstValue),
        occurrenceCount = 0;
      for (; matchStart >= 0 && occurrenceCount++ < 3;) {
        const pattern = {
          ...basePattern,
          prefix: value.slice(0, matchStart),
          suffix: value.slice(matchStart + firstValue.length),
        };
        if (
          pattern.prefix.length <= 1e3 &&
          pattern.suffix.length <= 1e3 &&
          exampleRows.every(
            (exampleRecord, valueIndex) =>
              pattern.prefix + patternValues[valueIndex] + pattern.suffix ===
              exampleRecord.example.value,
          )
        ) {
          const patternKey = JSON.stringify(pattern);
          if (
            !seenPatterns.has(patternKey) &&
            (seenPatterns.add(patternKey),
            candidates.push({
              id: `pattern-${candidates.length + 1}`,
              pattern: pattern,
              description: describeInferredTextRule(pattern),
              matchedExamples: examples.length,
            }),
            candidates.length >= textInferenceLimits.inferenceCandidates)
          ) {
            truncated = true;
            break e;
          }
        }
        matchStart = value.indexOf(firstValue, matchStart + 1);
      }
    }
  return { candidates: candidates, searched: searched, truncated: truncated };
}
const inferCellValueType = (value) =>
    value === null
      ? "blank"
      : typeof value == "number"
        ? "number"
        : typeof value == "boolean"
          ? "boolean"
          : "text",
  uniqueSourceReferences = (sourcedValues) => [
    ...new Set(sourcedValues.flatMap((sourcedValue) => sourcedValue.sources)),
  ],
  createSourcedValue = (
    value,
    sources,
    valueType = inferCellValueType(value),
  ) => ({
    value: value,
    type: valueType,
    sources: sources,
    exception: false,
  });
function assertRecipeFieldInRange(fieldIndexes, columnCount) {
  if (fieldIndexes.some((fieldIndex) => fieldIndex >= columnCount))
    throw Error(
      `A recipe field exceeds the current ${columnCount} columns. Review steps in order.`,
    );
}
function splitTextWithQuoteRules(text, rule) {
  if (rule.mode === "fixed") {
    const fieldCharacters = [...text],
      boundaries = [0, ...rule.breaks, fieldCharacters.length];
    return boundaries
      .slice(0, -1)
      .map((segmentStart, segmentIndex) =>
        fieldCharacters
          .slice(
            segmentStart,
            Math.max(segmentStart, boundaries[segmentIndex + 1]),
          )
          .join(""),
      );
  }
  const delimiterSet = new Set([...rule.delimiters]),
    characters = [...text],
    fields = [];
  let fieldText = "",
    isInQuotes = false,
    hasClosedQuote = false,
    hasFieldStarted = false,
    isAfterDelimiter = false;
  const finishField = () => {
    if (
      (fields.push(fieldText),
      (fieldText = ""),
      (hasClosedQuote = false),
      (hasFieldStarted = false),
      fields.length >= 50)
    )
      throw Error("More than 50 split fields.");
  };
  for (let charIndex = 0; charIndex < characters.length; charIndex++) {
    const character = characters[charIndex];
    if (isInQuotes)
      (character !== '"'
        ? (fieldText += character)
        : characters[charIndex + 1] === '"'
          ? ((fieldText += '"'), charIndex++)
          : ((isInQuotes = false), (hasClosedQuote = true)),
        (isAfterDelimiter = false));
    else if (delimiterSet.has(character))
      ((rule.collapse && isAfterDelimiter) || finishField(),
        (isAfterDelimiter = true));
    else {
      if (hasClosedQuote) throw Error("Unexpected text after a closing quote.");
      if (character === '"') {
        if (hasFieldStarted) throw Error("A quote must begin a field.");
        ((isInQuotes = true), (hasFieldStarted = true));
      } else ((fieldText += character), (hasFieldStarted = true));
      isAfterDelimiter = false;
    }
  }
  if (isInQuotes) throw Error("The quoted field is not closed.");
  return (fields.push(fieldText), fields);
}
function convertFieldToTypedValue(value, conversion, dateSystem = "1900") {
  if (conversion.to === "text") return coerceCellValueToText(value);
  if (isNullOrEmptyValue(value)) return null;
  if (conversion.to === "boolean") {
    if (typeof value == "boolean") return value;
    if (typeof value == "string" && /^(true|false)$/i.test(value.trim()))
      return value.trim().toUpperCase() === "TRUE";
    throw Error("Expected TRUE or FALSE.");
  }
  if (conversion.to === "date") {
    if (
      typeof value == "number" &&
      Number.isInteger(value) &&
      value >= (dateSystem === "1904" ? 0 : 1) &&
      value <= dateSystemMaxSerial(dateSystem)
    )
      return value;
    if (
      typeof value != "string" ||
      !/^\s*(?:\d{4}-\d{1,2}-\d{1,2}|\d{1,2}\/\d{1,2}\/\d{4})\s*$/.test(value)
    )
      throw Error(
        "Expected a full-year ISO date or an explicitly ordered slash date.",
      );
    const dateText = parseDateText(value, {
      textDateOrder: conversion.dateOrder,
      dateSystem: dateSystem,
    });
    if (dateText === undefined)
      throw Error(`The date is impossible or outside ${dateSystem}–9999.`);
    return dateText;
  }
  if (typeof value == "boolean")
    throw Error("Boolean values are not numeric fields.");
  let numericValue;
  if (typeof value == "number") numericValue = value;
  else {
    let numericText = value.trim();
    /^\(.*\)$/.test(numericText) &&
      (numericText = "-" + numericText.slice(1, -1));
    const exponentSuffix = /[eE][+-]?\d+$/.exec(numericText)?.[0] ?? "",
      decimalParts = (
        exponentSuffix
          ? numericText.slice(0, -exponentSuffix.length)
          : numericText
      ).split(conversion.decimal);
    if (decimalParts.length > 2)
      throw Error("More than one decimal separator.");
    const sign = /^[+-]/.exec(decimalParts[0])?.[0] ?? "",
      integerText = decimalParts[0].slice(sign.length),
      fractionText = decimalParts[1];
    if (conversion.group && integerText.includes(conversion.group)) {
      const groups = integerText.split(conversion.group);
      if (
        !/^\d{1,3}$/.test(groups[0]) ||
        groups.slice(1).some((group) => !/^\d{3}$/.test(group))
      )
        throw Error(
          "Grouping separators must separate groups of three digits.",
        );
    }
    const digits = conversion.group
      ? integerText.split(conversion.group).join("")
      : integerText;
    if (
      !/^\d*$/.test(digits) ||
      (fractionText !== undefined && !/^\d+$/.test(fractionText)) ||
      (!digits && !fractionText)
    )
      throw Error("The field does not match the explicit numeric separators.");
    if (
      (digits + (fractionText ?? "")).replace(/^0+/, "").replace(/0+$/, "")
        .length > 15
    )
      throw Error(
        "More than 15 significant digits; original text is retained.",
      );
    const normalizedText =
      sign +
      (digits || "0") +
      (fractionText === undefined ? "" : "." + fractionText) +
      exponentSuffix;
    if (
      ((numericValue = Number(normalizedText)),
      !Number.isFinite(numericValue) ||
        (numericValue === 0 && /[1-9]/.test(digits + (fractionText ?? ""))) ||
        (Number.isInteger(numericValue) && !Number.isSafeInteger(numericValue)))
    )
      throw Error("Numeric conversion exceeds supported precision or range.");
  }
  if (!Number.isFinite(numericValue)) throw Error("Expected a finite number.");
  if (conversion.to === "integer" && !Number.isSafeInteger(numericValue))
    throw Error("Expected a safe whole number.");
  return numericValue;
}
function prepareTransformationInput(workbook, recipe, permissions) {
  const request = validateFieldTransformRecipe(recipe),
    context = validateTextToColumnsPermissions(workbook, permissions),
    sheet = workbook.sheets.find((sheet) => sheet.id === request.sourceSheetId);
  if (!sheet) throw Error("Choose an existing source worksheet.");
  if (!context.allowNewSheet || workbook.structureProtection?.enabled)
    throw Error("Creating a derived worksheet is not permitted.");
  if (
    workbook.sheets.some(
      (sheet) => sheet.name.toLowerCase() === request.outputName.toLowerCase(),
    )
  )
    throw Error("The output worksheet name is already used.");
  const rangeAddress = parseRangeAddress(request.sourceRange),
    columnCount = rangeAddress.lastCol - rangeAddress.firstCol + 1,
    rowCount = rangeAddress.lastRow - rangeAddress.firstRow + 1;
  if (
    ((request.sourceRange = `${formatCellAddress(rangeAddress.firstRow, rangeAddress.firstCol)}:${formatCellAddress(rangeAddress.lastRow, rangeAddress.lastCol)}`),
    columnCount > 50 || columnCount * rowCount > 500)
  )
    throw Error("Select at most 500 source cells and 50 fields.");
  if (rowCount <= Number(request.header))
    throw Error("Select at least one data row in addition to the heading row.");
  if (
    (assertSelectionInAuthorizedSheet(sheet, request.sourceRange, context),
    sheet.merges?.some((merge) => {
      const mergeRange = parseRangeAddress(merge);
      return (
        mergeRange.firstRow <= rangeAddress.lastRow &&
        mergeRange.lastRow >= rangeAddress.firstRow &&
        mergeRange.firstCol <= rangeAddress.lastCol &&
        mergeRange.lastCol >= rangeAddress.firstCol
      );
    }))
  )
    throw Error("Unmerge the selected source before reshaping fields.");
  (assertOutputNameUnreferenced(workbook, request.outputName),
    discoverArrayFormulaAnchors(
      workbook,
      [{ sheet: sheet, range: request.sourceRange }],
      context,
    ));
  const sources = [],
    rows = [];
  let sourceTextLength = 0;
  for (
    let firstRow = rangeAddress.firstRow;
    firstRow <= rangeAddress.lastRow;
    firstRow++
  ) {
    const fields = [];
    for (
      let firstCol = rangeAddress.firstCol;
      firstCol <= rangeAddress.lastCol;
      firstCol++
    ) {
      const cellAddress = formatCellAddress(firstRow, firstCol),
        cell = sheet.cells[cellAddress];
      if (
        cell &&
        typeof cell.raw == "string" &&
        cell.raw.startsWith("=") &&
        !cell.literal
      )
        throw Error(
          `Source ${cellAddress} is a formula. Create a reviewed values snapshot first.`,
        );
      if (
        (sourceTextLength += coerceCellValueToText(cell?.raw ?? null).length) >
        2e5
      )
        throw Error(
          "Source text exceeds the 200,000-character processing budget.",
        );
      const rowOrigin = findRowOrigin(workbook, sheet.id, firstRow);
      (sources.push({
        address: cellAddress,
        ...(cell ? { original: structuredClone(cell) } : {}),
        ...(rowOrigin ? { origin: rowOrigin } : {}),
      }),
        fields.push(createSourcedValue(cell?.raw ?? null, [cellAddress])));
    }
    rows.push({ row: firstRow + 1, fields: fields });
  }
  const headers = request.header
    ? rows
        .shift()
        .fields.map(
          (field, columnIndex) =>
            coerceCellValueToText(field.value) || `Field ${columnIndex + 1}`,
        )
    : Array.from(
        { length: columnCount },
        (_, fieldIndex) => `Field ${fieldIndex + 1}`,
      );
  if (headers.some((header) => header.length > 100))
    throw Error("Source headings must be at most 100 characters.");
  return {
    request: request,
    context: context,
    source: sheet,
    sources: sources,
    rows: rows,
    headers: headers,
  };
}
function applyTransformationSteps(rows, headers, steps, dateSystem = "1900") {
  const exceptions = [],
    stages = [],
    recordException = (inputValue, row, field, step, parseError) => {
      const reason =
        parseError instanceof Error
          ? parseError.message
          : "The transformation could not parse this value.";
      return (
        exceptions.push({
          step: step,
          row: row,
          field: field,
          raw: inputValue.value,
          sources: [...inputValue.sources],
          reason: reason,
        }),
        { ...inputValue, exception: true }
      );
    };
  return (
    steps.forEach((transformStep, stepIndex) => {
      const step = stepIndex + 1,
        rowsSnapshot = JSON.stringify(rows),
        exceptionCountBefore = exceptions.length,
        stepColumns =
          "columns" in transformStep
            ? transformStep.columns
            : "column" in transformStep
              ? [transformStep.column]
              : transformStep.pattern.parts.map((part) => part.column);
      if (
        (assertRecipeFieldInRange(stepColumns, headers.length),
        transformStep.kind === "split")
      ) {
        let splitFieldCount =
          transformStep.mode === "fixed" ? transformStep.breaks.length + 1 : 1;
        const splitResults = rows.map((row) => {
          const sourceField = row.fields[transformStep.column];
          if (sourceField.exception) return [sourceField];
          try {
            const pieces = splitTextWithQuoteRules(
              coerceCellValueToText(sourceField.value),
              transformStep,
            );
            return (
              (splitFieldCount = Math.max(splitFieldCount, pieces.length)),
              pieces.map((piece) =>
                createSourcedValue(piece, [...sourceField.sources]),
              )
            );
          } catch (error) {
            return [
              recordException(
                sourceField,
                row.row,
                headers[transformStep.column],
                step,
                error,
              ),
            ];
          }
        });
        if (headers.length - 1 + splitFieldCount > 50)
          throw Error("The split would exceed 50 fields.");
        rows.forEach((row, splitRowIndex) => {
          const originalField = row.fields[transformStep.column],
            replacementFields = splitResults[splitRowIndex];
          for (; replacementFields.length < splitFieldCount;)
            replacementFields.push({
              ...createSourcedValue("", [...originalField.sources]),
              exception: replacementFields[0].exception,
            });
          row.fields.splice(transformStep.column, 1, ...replacementFields);
        });
        const originalHeader = headers[transformStep.column];
        headers.splice(
          transformStep.column,
          1,
          ...Array.from(
            { length: splitFieldCount },
            (_, splitIndex) =>
              `${originalHeader.slice(0, 130)}.${splitIndex + 1}`,
          ),
        );
      } else if (transformStep.kind === "concat") {
        if (headers.length >= 50)
          throw Error("Concatenation would exceed 50 fields.");
        for (const row of rows) {
          const concatInputs = transformStep.columns.map(
              (column) => row.fields[column],
            ),
            sourcedValue = createSourcedValue(
              concatInputs
                .filter(
                  (candidateField) =>
                    !transformStep.skipBlank ||
                    !isNullOrEmptyValue(candidateField.value),
                )
                .map((joinedField) => coerceCellValueToText(joinedField.value))
                .join(transformStep.separator),
              uniqueSourceReferences(concatInputs),
            );
          ((sourcedValue.exception = concatInputs.some(
            (failedField) => failedField.exception,
          )),
            row.fields.push(sourcedValue));
        }
        headers.push(transformStep.name);
      } else if (transformStep.kind === "select") {
        for (const row of rows)
          row.fields = transformStep.columns.map(
            (column) => row.fields[column],
          );
        headers.splice(
          0,
          headers.length,
          ...transformStep.columns.map((column) => headers[column]),
        );
      } else if (transformStep.kind === "fill") {
        assertRecipeFieldInRange(transformStep.groupColumns, headers.length);
        const orderedRows =
          transformStep.direction === "up" ? [...rows].reverse() : rows;
        for (const column of transformStep.columns) {
          let lastFilledField,
            previousGroupKey = "";
          for (const orderedRow of orderedRows) {
            const groupFields = transformStep.groupColumns.map(
                (groupColumn) => orderedRow.fields[groupColumn],
              ),
              groupKey = JSON.stringify(
                groupFields.map((groupField) => groupField.value),
              );
            (groupKey !== previousGroupKey ||
              groupFields.some(
                (checkedGroupField) => checkedGroupField.exception,
              )) &&
              ((lastFilledField = undefined), (previousGroupKey = groupKey));
            const currentField = orderedRow.fields[column];
            if (currentField.exception) {
              lastFilledField = undefined;
              continue;
            }
            if (!isNullOrEmptyValue(currentField.value)) {
              lastFilledField = currentField;
              continue;
            }
            transformStep.direction === "value"
              ? (orderedRow.fields[column] = createSourcedValue(
                  transformStep.value,
                  [...currentField.sources],
                ))
              : lastFilledField &&
                (orderedRow.fields[column] = {
                  ...lastFilledField,
                  sources: uniqueSourceReferences([
                    currentField,
                    lastFilledField,
                    ...groupFields,
                  ]),
                });
          }
        }
      } else if (transformStep.kind === "pattern") {
        if (
          (assertRecipeFieldInRange(
            transformStep.inferenceColumns,
            headers.length,
          ),
          transformStep.pattern.parts.some(
            (part) => !transformStep.inferenceColumns.includes(part.column),
          ))
        )
          throw Error(
            "The inferred pattern reads a field outside its approved inference columns.",
          );
        if (headers.length >= 50)
          throw Error("The inferred output would exceed 50 fields.");
        const seenExampleRows = new Set();
        for (const example of transformStep.examples) {
          const row = rows.find((row) => row.row === example.row);
          if (!row || seenExampleRows.has(example.row))
            throw Error(
              "An inference example is outside the data rows or duplicated.",
            );
          if (
            (seenExampleRows.add(example.row),
            transformStep.pattern.parts.some(
              (part) => row.fields[part.column].exception,
            ) ||
              applyInferredTextPattern(
                row.fields.map((field) => field.value),
                transformStep.pattern,
              ) !== example.value)
          )
            throw Error(
              "The inferred pattern no longer matches every reviewed example.",
            );
        }
        for (const row of rows) {
          const inputFields = transformStep.pattern.parts.map(
              (part) => row.fields[part.column],
            ),
            patternField = {
              ...inputFields[0],
              sources: uniqueSourceReferences(inputFields),
            };
          try {
            if (
              inputFields.some(
                (patternInputField) => patternInputField.exception,
              )
            )
              throw Error("An inferred input has an earlier parse exception.");
            row.fields.push(
              createSourcedValue(
                applyInferredTextPattern(
                  row.fields.map((field) => field.value),
                  transformStep.pattern,
                ),
                patternField.sources,
              ),
            );
          } catch (error) {
            row.fields.push(
              recordException(
                patternField,
                row.row,
                transformStep.name,
                step,
                error,
              ),
            );
          }
        }
        headers.push(transformStep.name);
      } else
        for (const row of rows)
          for (const column of transformStep.columns) {
            const targetField = row.fields[column];
            if (!targetField.exception)
              try {
                let value = targetField.value,
                  type = targetField.type;
                if (transformStep.kind === "convert")
                  ((value = convertFieldToTypedValue(
                    targetField.value,
                    transformStep,
                    dateSystem,
                  )),
                    (type =
                      value === null
                        ? "blank"
                        : transformStep.to === "date"
                          ? "date"
                          : inferCellValueType(value)));
                else if (!isNullOrEmptyValue(targetField.value)) {
                  const sourceText = coerceCellValueToText(targetField.value);
                  ((value =
                    transformStep.kind === "trim"
                      ? transformStep.mode === "collapse"
                        ? sourceText.trim().replace(/\s+/gu, " ")
                        : sourceText.trim()
                      : transformStep.kind === "clean"
                        ? sourceText.replace(/[\u0000-\u001f]/g, "")
                        : transformStep.kind === "case"
                          ? applyTextCasing(sourceText, transformStep.mode)
                          : sourceText
                              .split(transformStep.find)
                              .join(transformStep.replacement)),
                    (type = "text"));
                }
                if (typeof value == "string" && value.length > 32767)
                  throw Error(
                    "The transformed value exceeds the cell text limit.",
                  );
                row.fields[column] = {
                  ...targetField,
                  value: value,
                  type: type,
                };
              } catch (error) {
                row.fields[column] = recordException(
                  targetField,
                  row.row,
                  headers[column],
                  step,
                  error,
                );
              }
          }
      if (rows.length * headers.length > 5e3)
        throw Error("The intermediate result exceeds 5,000 cells.");
      let intermediateTextLength = 0;
      for (const row of rows)
        for (const field of row.fields) {
          if (typeof field.value == "string" && field.value.length > 32767)
            throw Error("A result exceeds the 32,767-character cell limit.");
          if (
            (intermediateTextLength += coerceCellValueToText(
              field.value,
            ).length) > 1e6
          )
            throw Error(
              "Intermediate results exceed the 1,000,000-character limit.",
            );
        }
      const previousRows = JSON.parse(rowsSnapshot);
      let changedCells = 0;
      (rows.forEach((row, rowIndex) =>
        row.fields.forEach((field, fieldIndex) => {
          JSON.stringify(field) !==
            JSON.stringify(previousRows[rowIndex].fields[fieldIndex]) &&
            changedCells++;
        }),
      ),
        stages.push({
          step: step,
          kind: transformStep.kind,
          fields: [...headers],
          changedCells: changedCells,
          exceptions: exceptions.length - exceptionCountBefore,
        }));
    }),
    { exceptions: exceptions, stages: stages }
  );
}
function createAnnotatedCellValue(raw, note, valueType) {
  return normalizeCellRecordInput({
    raw: raw,
    note: note,
    ...(valueType === "date" ? { format: "date" } : {}),
    ...(typeof raw == "string"
      ? { numFmt: "@", ...(raw.startsWith("=") ? { literal: true } : {}) }
      : {}),
  });
}
function buildTransformationPreview(workbook, recipe, permissions = {}) {
  const { request, context, source, sources, rows, headers } =
    prepareTransformationInput(workbook, recipe, permissions);
  if (!request.steps.length)
    throw Error("Add at least one transformation step before previewing.");
  const { exceptions, stages } = applyTransformationSteps(
      rows,
      headers,
      request.steps,
      workbook.datePolicy?.dateSystem,
    ),
    outputs = [],
    cells = {};
  if (
    (headers.forEach((header, columnIndex) => {
      cells[formatCellAddress(0, columnIndex)] = {
        raw: header,
        literal: true,
        style: { bold: true },
      };
    }),
    rows.forEach((row, rowIndex) =>
      row.fields.forEach((field, fieldIndex) => {
        const cellAddress = formatCellAddress(rowIndex + 1, fieldIndex);
        outputs.push({
          address: cellAddress,
          row: row.row,
          field: headers[fieldIndex],
          ...field,
        });
        const note = `Clean & Reshape derivation
Workbook: ${JSON.stringify(workbook.id)}
Source revision: ${workbook.revision}
Worksheet: ${JSON.stringify(source.name)} (${JSON.stringify(source.id)})
Original cells: ${field.sources.join(", ")}
Result type: ${field.type}
${field.exception ? "Parse exception retained; see the exceptions appendix." : "Original source values are retained."}`;
        cells[cellAddress] = createAnnotatedCellValue(
          field.value,
          note,
          field.type,
        );
      }),
    ),
    exceptions.length)
  ) {
    const exceptionTitleRow = rows.length + 3;
    ((cells[formatCellAddress(exceptionTitleRow, 0)] = {
      raw: "Transformation exceptions",
      style: { bold: true },
    }),
      [
        "Source row",
        "Field",
        "Step",
        "Raw value",
        "Original cells",
        "Reason",
      ].forEach(
        (raw, headerColumnIndex) =>
          (cells[formatCellAddress(exceptionTitleRow + 1, headerColumnIndex)] =
            {
              raw: raw,
              style: { bold: true },
            }),
      ),
      exceptions.forEach((exception, exceptionIndex) =>
        [
          exception.row,
          exception.field,
          exception.step,
          exception.raw,
          exception.sources
            .map((sourceAddress) => `${source.name}!${sourceAddress}`)
            .join(", "),
          exception.reason,
        ].forEach(
          (raw, exceptionColumnIndex) =>
            (cells[
              formatCellAddress(
                exceptionTitleRow + 2 + exceptionIndex,
                exceptionColumnIndex,
              )
            ] = createAnnotatedCellValue(
              raw,
              `Retained parsing exception at source revision ${workbook.revision}`,
              inferCellValueType(raw),
            )),
        ),
      ));
  }
  if (Object.keys(cells).length > 5e3)
    throw Error("Results and exception appendix exceed 5,000 output cells.");
  if (
    workbook.sheets.reduce(
      (total, sheet) => total + Object.keys(sheet.cells).length,
      0,
    ) +
      Object.keys(cells).length >
    workbookLimits.maxCells
  )
    throw Error(
      "Derived output exceeds the 100,000 stored-cell workbook budget.",
    );
  const preview = {
    snapshot: JSON.stringify(workbook),
    context,
    request,
    sourceSheetName: source.name,
    sources,
    headers,
    outputs: outputs,
    exceptions,
    stages,
    cells: cells,
    targetRange: `A1:${formatCellAddress(exceptions.length ? rows.length + 4 + exceptions.length : rows.length, Math.max(headers.length, exceptions.length ? 6 : 0) - 1)}`,
    counts: {
      sourceRows: rows.length,
      sourceCells: sources.length,
      outputCells: Object.keys(cells).length,
      changedCells: outputs.filter(
        (output) =>
          output.sources.length !== 1 ||
          JSON.stringify(
            sources.find((source) => source.address === output.sources[0])
              ?.original?.raw ?? null,
          ) !== JSON.stringify(output.value),
      ).length,
      exceptions: exceptions.length,
      exceptionRows: new Set(exceptions.map((exception) => exception.row)).size,
    },
  };
  return (
    validateFieldTransformEvidence(
      [
        ...(workbook.fieldTransforms ?? []),
        buildTransformationRecord(
          workbook,
          preview,
          "preview-id",
          "preview-output",
          "2026-01-01T00:00:00.000Z",
        ),
      ],
      {},
    ),
    preview
  );
}
function buildTransformationRecord(
  workbook,
  preview,
  recordId,
  outputSheetId,
  createdAt,
) {
  return {
    id: recordId,
    workbookId: workbook.id,
    sourceRevision: workbook.revision,
    sourceSheetId: preview.request.sourceSheetId,
    sourceSheetName: preview.sourceSheetName,
    outputSheetId: outputSheetId,
    outputSheetName: preview.request.outputName,
    createdAt: createdAt,
    semantics: "reviewed-fields-v1",
    request: structuredClone(preview.request),
    sources: structuredClone(preview.sources),
    outputs: structuredClone(preview.outputs),
    exceptions: structuredClone(preview.exceptions),
  };
}
function commitApprovedTransformation(workbook, preview, approval) {
  if (approval.approved !== true)
    throw Error(
      "Review and approve the complete transformation preview first.",
    );
  if (preview.snapshot !== JSON.stringify(workbook))
    throw Error("The workbook changed. Preview the transformation again.");
  const transformationPreview = buildTransformationPreview(
    workbook,
    preview.request,
    approval,
  );
  if (JSON.stringify(transformationPreview) !== JSON.stringify(preview))
    throw Error("The preview or authorized context changed. Preview again.");
  if (preview.exceptions.length && approval.acceptExceptions !== true)
    throw Error("Acknowledge the retained parse exceptions before applying.");
  const nextWorkbook = structuredClone(workbook),
    worksheetRecord = createWorksheetRecord(preview.request.outputName);
  return (
    (worksheetRecord.cells = structuredClone(preview.cells)),
    nextWorkbook.sheets.push(worksheetRecord),
    (nextWorkbook.fieldTransforms = [
      ...(nextWorkbook.fieldTransforms ?? []),
      buildTransformationRecord(
        workbook,
        preview,
        crypto.randomUUID(),
        worksheetRecord.id,
        new Date().toISOString(),
      ),
    ]),
    normalizeWorkbookInput(nextWorkbook)
  );
}
function selectExceptionFreeRows(
  workbook,
  recipe,
  examples,
  columns,
  permissions = {},
) {
  const transformationInput = prepareTransformationInput(
    workbook,
    recipe,
    permissions,
  );
  applyTransformationSteps(
    transformationInput.rows,
    transformationInput.headers,
    transformationInput.request.steps,
    workbook.datePolicy?.dateSystem,
  );
  const exceptionFreeRows = transformationInput.rows.filter((row) =>
    columns.every(
      (column) => row.fields[column] && !row.fields[column].exception,
    ),
  );
  return inferTextPatternFromExamples(
    exceptionFreeRows.map((exceptionFreeRow) => ({
      row: exceptionFreeRow.row,
      values: exceptionFreeRow.fields.map((field) => field.value),
    })),
    columns,
    examples,
  );
}
function verifyTransformationPreview(workbook, preview, permissions = {}) {
  if (
    preview.snapshot !== JSON.stringify(workbook) ||
    JSON.stringify(
      buildTransformationPreview(workbook, preview.request, permissions),
    ) !== JSON.stringify(preview)
  )
    throw Error(
      "The transformation preview is stale or altered. Preview again.",
    );
  const { snapshot, context, cells, ...previewDetails } = preview;
  return JSON.stringify(
    {
      schemaVersion: 1,
      kind: "field-transformation-preview",
      workbookId: workbook.id,
      sourceRevision: workbook.revision,
      ...previewDetails,
    },
    null,
    2,
  );
}
const expressionFunctionArity = {
    ABS: [1, 1],
    ROUND: [1, 2],
    UPPER: [1, 1],
    LOWER: [1, 1],
    TRIM: [1, 1],
    LEN: [1, 1],
    CONCAT: [1, 20],
    COALESCE: [1, 20],
    IF: [3, 3],
    AND: [1, 20],
    OR: [1, 20],
    NOT: [1, 1],
  },
  assertFiniteNumericOperand = (operand) => {
    if (typeof operand != "number" || !Number.isFinite(operand))
      throw Error(
        "This expression requires a finite number. Add an explicit Types step.",
      );
    return operand;
  },
  assertTextOperand = (operand) => {
    if (typeof operand != "string")
      throw Error("This expression requires text. Add an explicit Types step.");
    return operand;
  },
  assertBooleanCondition = (condition) => {
    if (typeof condition != "boolean")
      throw Error("IF/AND/OR/NOT conditions must be Boolean comparisons.");
    return condition;
  },
  assertDerivedValueLimits = (derivedValue) => {
    if (
      (typeof derivedValue == "number" && !Number.isFinite(derivedValue)) ||
      (typeof derivedValue == "string" && derivedValue.length > 32767)
    )
      throw Error("Derived value exceeds finite numeric or text limits.");
    return derivedValue;
  };
function parseRowExpression(expression, fields) {
  if (
    typeof expression != "string" ||
    !expression.trim() ||
    expression.length > 1e3
  )
    throw Error("Use a row expression of 1–1,000 characters.");
  let cursor = 0,
    nodeCount = 0,
    nestingDepth = 0;
  const expressionText = expression.trim().replace(/^=/, ""),
    skipWhitespace = () => {
      for (
        ;
        /\s/.test(expressionText[cursor] ?? "") &&
        cursor < expressionText.length;
      )
        cursor++;
    },
    peekChar = () => (skipWhitespace(), expressionText[cursor]),
    throwSyntaxError = () => {
      throw Error(
        `Invalid row expression near character ${cursor + 1}. Use [Field], literals, operators and supported functions.`,
      );
    },
    countNode = (node) => {
      if (++nodeCount > 200)
        throw Error("Row expressions are limited to 200 nodes.");
      return node;
    };
  function parsePrimary() {
    if (++nestingDepth > 32)
      throw Error("Row expression nesting exceeds 32 levels.");
    try {
      let operator = peekChar();
      if (operator === "+" || operator === "-")
        return (
          cursor++,
          countNode({
            kind: "unary",
            operator: operator,
            value: parsePrimary(),
          })
        );
      if (operator === "(") {
        cursor++;
        const groupedNode = parseBinary(0);
        return peekChar() !== ")"
          ? throwSyntaxError()
          : (cursor++, groupedNode);
      }
      if (operator === '"') {
        const stringMatch = expressionText
          .slice(cursor)
          .match(/^"(?:[^"\\]|\\.)*"/);
        if (!stringMatch) return throwSyntaxError();
        cursor += stringMatch[0].length;
        try {
          return countNode({
            kind: "value",
            value: JSON.parse(stringMatch[0]),
          });
        } catch {
          return throwSyntaxError();
        }
      }
      if (operator === "[") {
        cursor++;
        let fieldName = "",
          isFieldClosed = false;
        for (; cursor < expressionText.length;) {
          if (expressionText[cursor] === "]") {
            if (expressionText[cursor + 1] === "]") {
              ((fieldName += "]"), (cursor += 2));
              continue;
            }
            (cursor++, (isFieldClosed = true));
            break;
          }
          fieldName += expressionText[cursor++];
        }
        const fieldIndex = fields.indexOf(fieldName);
        if (!isFieldClosed || fieldIndex < 0)
          throw Error(`Unknown row-expression field [${fieldName}].`);
        return countNode({ kind: "field", index: fieldIndex });
      }
      const numberMatch = expressionText
        .slice(cursor)
        .match(/^(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?/);
      if (numberMatch)
        return (
          (cursor += numberMatch[0].length),
          countNode({
            kind: "value",
            value: assertDerivedValueLimits(Number(numberMatch[0])),
          })
        );
      const identifierMatch = expressionText
        .slice(cursor)
        .match(/^[A-Za-z_][A-Za-z0-9_]*/);
      if (!identifierMatch) return throwSyntaxError();
      cursor += identifierMatch[0].length;
      const identifierName = identifierMatch[0].toUpperCase();
      if (peekChar() !== "(") {
        if (
          identifierName === "TRUE" ||
          identifierName === "FALSE" ||
          identifierName === "NULL"
        )
          return countNode({
            kind: "value",
            value: identifierName === "NULL" ? null : identifierName === "TRUE",
          });
        throw Error(
          `Unknown identifier ${identifierName}. Column references use [Field].`,
        );
      }
      if (!Object.hasOwn(expressionFunctionArity, identifierName))
        throw Error(
          `Unsupported row function ${identifierName}. Use ${Object.keys(expressionFunctionArity).join(", ")}.`,
        );
      cursor++;
      const args = [];
      if (peekChar() !== ")")
        for (; args.push(parseBinary(0)), peekChar() === ",";) cursor++;
      if (peekChar() !== ")") return throwSyntaxError();
      cursor++;
      const [minArity, maxArity] = expressionFunctionArity[identifierName];
      if (args.length < minArity || args.length > maxArity)
        throw Error(
          `${identifierName} expects ${minArity === maxArity ? minArity : `${minArity}–${maxArity}`} arguments.`,
        );
      return countNode({ kind: "call", name: identifierName, args: args });
    } finally {
      nestingDepth--;
    }
  }
  const operatorPrecedence = {
    "=": 1,
    "<>": 1,
    "<": 1,
    ">": 1,
    "<=": 1,
    ">=": 1,
    "&": 2,
    "+": 3,
    "-": 3,
    "*": 4,
    "/": 4,
  };
  function parseBinary(minPrecedence) {
    let left = parsePrimary();
    for (;;) {
      skipWhitespace();
      const operator = expressionText
        .slice(cursor)
        .match(/^(<>|<=|>=|[=<>+*/&-])/)?.[0];
      if (!operator || operatorPrecedence[operator] < minPrecedence) break;
      ((cursor += operator.length),
        (left = countNode({
          kind: "binary",
          operator: operator,
          left: left,
          right: parseBinary(operatorPrecedence[operator] + 1),
        })));
    }
    return left;
  }
  const rootNode = parseBinary(0);
  if ((skipWhitespace(), cursor !== expressionText.length))
    return throwSyntaxError();
  function evaluateNode(node, rowValues) {
    if (node.kind === "value") return node.value;
    if (node.kind === "field") return rowValues[node.index];
    if (node.kind === "unary")
      return assertDerivedValueLimits(
        assertFiniteNumericOperand(evaluateNode(node.value, rowValues)) *
          (node.operator === "-" ? -1 : 1),
      );
    if (node.kind === "binary") {
      const leftValue = evaluateNode(node.left, rowValues),
        rightValue = evaluateNode(node.right, rowValues),
        operator = node.operator;
      if (operator === "=") return leftValue === rightValue;
      if (operator === "<>") return leftValue !== rightValue;
      if (["<", ">", "<=", ">="].includes(operator)) {
        if (
          typeof leftValue != typeof rightValue ||
          leftValue === null ||
          rightValue === null ||
          typeof leftValue == "boolean"
        )
          throw Error(
            "Comparisons require two numbers or two text values of the same type.",
          );
        return operator === "<"
          ? leftValue < rightValue
          : operator === ">"
            ? leftValue > rightValue
            : operator === "<="
              ? leftValue <= rightValue
              : leftValue >= rightValue;
      }
      if (operator === "&")
        return assertDerivedValueLimits(
          assertTextOperand(leftValue) + assertTextOperand(rightValue),
        );
      const leftNumber = assertFiniteNumericOperand(leftValue),
        rightNumber = assertFiniteNumericOperand(rightValue);
      if (operator === "/" && rightNumber === 0)
        throw Error("Derived expression divides by zero.");
      return assertDerivedValueLimits(
        operator === "+"
          ? leftNumber + rightNumber
          : operator === "-"
            ? leftNumber - rightNumber
            : operator === "*"
              ? leftNumber * rightNumber
              : leftNumber / rightNumber,
      );
    }
    if (node.name === "IF")
      return evaluateNode(
        node.args[
          assertBooleanCondition(evaluateNode(node.args[0], rowValues)) ? 1 : 2
        ],
        rowValues,
      );
    if (node.name === "COALESCE") {
      for (const arg of node.args) {
        const candidateValue = evaluateNode(arg, rowValues);
        if (candidateValue !== null && candidateValue !== "")
          return candidateValue;
      }
      return null;
    }
    const argValues = node.args.map((arg) => evaluateNode(arg, rowValues)),
      firstArgValue = argValues[0];
    switch (node.name) {
      case "ABS":
        return Math.abs(assertFiniteNumericOperand(firstArgValue));
      case "ROUND": {
        const digits =
          argValues.length === 1 ? 0 : assertFiniteNumericOperand(argValues[1]);
        if (!Number.isInteger(digits) || Math.abs(digits) > 12)
          throw Error("ROUND digits must be an integer from -12 to 12.");
        const numberToRound = assertFiniteNumericOperand(firstArgValue),
          scaleFactor = 10 ** digits;
        return assertDerivedValueLimits(
          (Math.sign(numberToRound) *
            Math.round(Math.abs(numberToRound) * scaleFactor)) /
            scaleFactor,
        );
      }
      case "UPPER":
        return assertDerivedValueLimits(
          assertTextOperand(firstArgValue).toLocaleUpperCase("en-US"),
        );
      case "LOWER":
        return assertDerivedValueLimits(
          assertTextOperand(firstArgValue).toLocaleLowerCase("en-US"),
        );
      case "TRIM":
        return assertTextOperand(firstArgValue).trim().replace(/\s+/gu, " ");
      case "LEN":
        return Array.from(assertTextOperand(firstArgValue)).length;
      case "CONCAT":
        return assertDerivedValueLimits(
          argValues.map(assertTextOperand).join(""),
        );
      case "AND":
        return argValues.map(assertBooleanCondition).every(Boolean);
      case "OR":
        return argValues.map(assertBooleanCondition).some(Boolean);
      case "NOT":
        return !assertBooleanCondition(firstArgValue);
      default:
        throw Error("Unknown row expression operation.");
    }
  }
  return (rowValues) =>
    assertDerivedValueLimits(evaluateNode(rootNode, rowValues));
}
const findFieldIndexOrThrow = (table, fieldName) => {
    const fieldIndex = table.fields.indexOf(fieldName);
    if (fieldIndex < 0)
      throw Error(
        `Field “${fieldName}” is missing from the step input schema.`,
      );
    return fieldIndex;
  },
  uniqueSortedSources = (rows) =>
    [...new Set(rows.flatMap((row) => row.sources))].sort(),
  profileFieldTypes = (table) =>
    table.fields.map((field, fieldIndex) => ({
      name: field,
      types: [
        ...new Set(
          table.rows.map((row) =>
            isNullOrEmptyValue(row.values[fieldIndex])
              ? "blank"
              : typeof row.values[fieldIndex],
          ),
        ),
      ].sort(),
      ...(Object.hasOwn(table.declaredTypes, field)
        ? { declaredType: table.declaredTypes[field] }
        : {}),
    }));
function buildFieldControlTotals(output) {
  return output.fields.map((field, fieldIndex) => {
    let sum = 0,
      compensation = 0,
      numericCount = 0,
      blankCount = 0;
    for (const row of output.rows) {
      const cellValue = row.values[fieldIndex];
      if (
        (isNullOrEmptyValue(cellValue) && blankCount++,
        typeof cellValue == "number")
      ) {
        numericCount++;
        const adjustedValue = cellValue - compensation,
          nextSum = sum + adjustedValue;
        if (
          ((compensation = nextSum - sum - adjustedValue),
          (sum = nextSum),
          !Number.isFinite(sum) || !Number.isFinite(compensation))
        )
          throw Error(
            `Control total for “${field}” exceeds finite numeric limits.`,
          );
      }
    }
    return {
      field: field,
      numericCount: numericCount,
      blankCount: blankCount,
      sum: sum,
    };
  });
}
function validateRecipeOutputTable(table) {
  if (
    !table.fields.length ||
    table.fields.length > transformationRecipeLimits.columns ||
    table.fields.some((field) => !field.trim() || field.length > 128) ||
    new Set(table.fields.map((field) => field.toLowerCase())).size !==
      table.fields.length
  )
    throw Error(
      "The result needs 1–50 distinct, nonblank column names of at most 128 characters.",
    );
  if (
    table.rows.length > transformationRecipeLimits.rows ||
    (table.rows.length + 1) * table.fields.length >
      transformationRecipeLimits.cells
  )
    throw Error(
      "A recipe step exceeds 1,000 rows or 20,000 cells. Narrow the source; no rows were truncated.",
    );
  let characterCount = 0,
    lineageLinkCount = 0;
  for (const row of table.rows) {
    if (row.values.length !== table.fields.length)
      throw Error("Invalid recipe row width.");
    lineageLinkCount += row.sources.length;
    for (const value of row.values)
      if (typeof value == "string") {
        if (((characterCount += value.length), value.length > 32767))
          throw Error("Recipe text exceeds the cell limit.");
      } else if (
        value !== null &&
        typeof value != "boolean" &&
        (typeof value != "number" || !Number.isFinite(value))
      )
        throw Error("Recipe output must contain finite scalar values.");
  }
  if (
    characterCount > transformationRecipeLimits.characters ||
    lineageLinkCount > transformationRecipeLimits.lineageLinks
  )
    throw Error(
      "Recipe text or source-lineage budget exceeded. Narrow the source; no evidence was truncated.",
    );
  return table;
}
function resolveRecipeSourceTable(workbook, source, hostContext) {
  const sheet = workbook.sheets.find((sheet) => sheet.id === source.sheetId);
  if (
    !sheet ||
    (hostContext.allowedSheetIds &&
      !hostContext.allowedSheetIds.includes(sheet.id))
  )
    throw Error("A recipe source is outside the authorized host context.");
  let range = source.range;
  if (source.tableId) {
    const sourceTable = sheet.tables?.find(
      (table) => table.id === source.tableId,
    );
    if (!sourceTable)
      throw Error(
        "The selected source table no longer exists. Review the recipe source.",
      );
    const rangeAddress = parseRangeAddress(sourceTable.range);
    range = `${columnIndexToLetters(rangeAddress.firstCol)}${rangeAddress.firstRow + 1}:${columnIndexToLetters(rangeAddress.lastCol)}${rangeAddress.lastRow + 1 - (sourceTable.showTotals ? 1 : 0)}`;
  }
  const sourceHeaders = readSourceHeaders(
      workbook,
      [{ sheetId: source.sheetId, range: range }],
      hostContext,
    ),
    snapshot = sourceHeaders.snapshots[0],
    output = validateRecipeOutputTable({
      fields: snapshot.headers,
      rows: snapshot.records.map((record) => ({
        values: record.values,
        sources: [record.locator],
      })),
      declaredTypes: {},
    }),
    evidence = {
      source: source,
      resolvedRange: snapshot.source.range,
      sheetName: sheet.name,
      hash: fingerprintCanonicalValue({
        workbookId: workbook.id,
        source: source,
        snapshot: snapshot,
        dependencies: sourceHeaders.dependencyCells,
        datePolicy: normalizeDatePolicy(workbook.datePolicy),
      }),
      schema: profileFieldTypes(output),
      rows: output.rows.length,
      controls: buildFieldControlTotals(output),
    };
  return { data: output, evidence: evidence };
}
function aggregateFieldValues(rows, fieldIndex, operation) {
  const nonBlankValues = rows
    .map((row) => row.values[fieldIndex])
    .filter((cellValue) => !isNullOrEmptyValue(cellValue));
  if (operation === "count") return nonBlankValues.length;
  if (operation === "distinctCount")
    return new Set(
      nonBlankValues.map((distinctValue) =>
        serializeTypedKeyValues([distinctValue]),
      ),
    ).size;
  if (nonBlankValues.some((nonBlankValue) => typeof nonBlankValue != "number"))
    throw Error(
      "Numeric aggregation encountered text or Boolean values. Add a Types step before aggregating.",
    );
  const numericValues = nonBlankValues;
  if (!numericValues.length) return operation === "sum" ? 0 : null;
  let result =
    operation === "min"
      ? Math.min(...numericValues)
      : operation === "max"
        ? Math.max(...numericValues)
        : numericValues.reduce((total, value) => total + value, 0);
  if (
    (operation === "average" && (result /= numericValues.length),
    !Number.isFinite(result))
  )
    throw Error("Aggregation exceeds finite numeric bounds.");
  return result;
}
function groupRowsByFields(table, keys) {
  const indices = keys.map((key) => findFieldIndexOrThrow(table, key)),
    groups = new Map();
  for (const row of table.rows) {
    const groupKey = serializeTypedKeyValues(
      indices.map((indice) => row.values[indice]),
    );
    groups.set(groupKey, [...(groups.get(groupKey) ?? []), row]);
  }
  return { indices: indices, groups: groups };
}
function executeRecipeStep(workbook, output, step, loadSourceTable) {
  const inputTable = output,
    resolveFieldIndices = (fieldNames) =>
      fieldNames.map((fieldName) =>
        findFieldIndexOrThrow(inputTable, fieldName),
      );
  switch (step.kind) {
    case "select": {
      const selectedIndices = resolveFieldIndices(step.fields);
      return {
        data: {
          fields: step.fields,
          rows: inputTable.rows.map((row) => ({
            ...row,
            values: selectedIndices.map(
              (selectedColumnIndex) => row.values[selectedColumnIndex],
            ),
          })),
          declaredTypes: Object.fromEntries(
            step.fields.flatMap((field) =>
              Object.hasOwn(inputTable.declaredTypes, field)
                ? [[field, inputTable.declaredTypes[field]]]
                : [],
            ),
          ),
        },
      };
    }
    case "rename": {
      if (
        (resolveFieldIndices(step.mappings.map((mapping) => mapping.from)),
        new Set(step.mappings.map((mapping) => mapping.from)).size !==
          step.mappings.length)
      )
        throw Error("Rename each input field only once per step.");
      const renameField = (originalName) =>
        step.mappings.find((mapping) => mapping.from === originalName)?.to ??
        originalName;
      return {
        data: {
          ...inputTable,
          fields: inputTable.fields.map(renameField),
          declaredTypes: Object.fromEntries(
            Object.entries(inputTable.declaredTypes).map(
              ([declaredFieldName, declaredType]) => [
                renameField(declaredFieldName),
                declaredType,
              ],
            ),
          ),
        },
      };
    }
    case "types": {
      const columns = resolveFieldIndices(step.fields);
      if (step.group && step.group === step.decimal)
        throw Error("Decimal and grouping separators must differ.");
      return {
        data: {
          ...inputTable,
          declaredTypes: {
            ...inputTable.declaredTypes,
            ...Object.fromEntries(step.fields.map((field) => [field, step.to])),
          },
          rows: inputTable.rows.map((row, rowIndex) => {
            try {
              return {
                ...row,
                values: row.values.map((value, columnIndex) =>
                  columns.includes(columnIndex)
                    ? convertFieldToTypedValue(
                        value,
                        {
                          kind: "convert",
                          columns: columns,
                          to: step.to,
                          decimal: step.decimal,
                          group: step.group,
                          dateOrder: step.dateOrder,
                        },
                        normalizeDatePolicy(workbook.datePolicy).dateSystem,
                      )
                    : value,
                ),
              };
            } catch (error) {
              throw Error(
                `Input record ${rowIndex + 1} (${row.sources.join(", ")}): ${error.message}`,
              );
            }
          }),
        },
      };
    }
    case "replace": {
      const replaceColumnIndices = resolveFieldIndices(step.fields);
      if (
        step.mode === "text" &&
        (typeof step.find != "string" ||
          !step.find ||
          typeof step.replacement != "string")
      )
        throw Error(
          "Text replace needs nonempty text to find and a text replacement.",
        );
      const replaceText = (cellText) => {
        const textParts = cellText.split(step.find);
        if (
          cellText.length +
            (textParts.length - 1) *
              (step.replacement.length - step.find.length) >
          32767
        )
          throw Error("Replacement exceeds the cell text limit.");
        return textParts.join(step.replacement);
      };
      return {
        data: {
          ...inputTable,
          declaredTypes: {},
          rows: inputTable.rows.map((row) => ({
            ...row,
            values: row.values.map((value, valueColumnIndex) =>
              replaceColumnIndices.includes(valueColumnIndex)
                ? step.mode === "value"
                  ? value === step.find
                    ? step.replacement
                    : value
                  : typeof value == "string"
                    ? replaceText(value)
                    : value
                : value,
            ),
          })),
        },
      };
    }
    case "derive": {
      if (
        inputTable.fields.some(
          (field) => field.toLowerCase() === step.name.toLowerCase(),
        )
      )
        throw Error("The derived column name already exists.");
      const rowExpression = parseRowExpression(
        step.expression,
        inputTable.fields,
      );
      return {
        data: {
          ...inputTable,
          fields: [...inputTable.fields, step.name],
          rows: inputTable.rows.map((row, deriveRowIndex) => {
            try {
              return {
                ...row,
                values: [...row.values, rowExpression(row.values)],
              };
            } catch (error) {
              throw Error(
                `Input record ${deriveRowIndex + 1} (${row.sources.join(", ")}): ${error.message}`,
              );
            }
          }),
        },
      };
    }
    case "filter": {
      const resolvedConditions = step.conditions.map((condition) => {
          const fieldIndexOrThrow = findFieldIndexOrThrow(
            inputTable,
            condition.field,
          );
          if (
            !["blank", "notBlank"].includes(condition.operator) &&
            condition.value === undefined
          )
            throw Error("This filter requires an explicit typed value.");
          if (
            ["contains", "startsWith", "endsWith"].includes(
              condition.operator,
            ) &&
            typeof condition.value != "string"
          )
            throw Error("Text filters require a text criterion.");
          if (
            ["gt", "gte", "lt", "lte"].includes(condition.operator) &&
            typeof condition.value != "number"
          )
            throw Error("Ordered filters require a numeric criterion.");
          return { ...condition, i: fieldIndexOrThrow };
        }),
        rows = inputTable.rows.filter((row, filterRowIndex) => {
          const conditionResults = resolvedConditions.map(
            (resolvedCondition) => {
              const cellValue = row.values[resolvedCondition.i],
                value = resolvedCondition.value;
              switch (resolvedCondition.operator) {
                case "blank":
                  return isNullOrEmptyValue(cellValue);
                case "notBlank":
                  return !isNullOrEmptyValue(cellValue);
                case "equals":
                  return cellValue === value;
                case "notEquals":
                  return cellValue !== value;
                case "contains":
                  return (
                    typeof cellValue == "string" && cellValue.includes(value)
                  );
                case "startsWith":
                  return (
                    typeof cellValue == "string" && cellValue.startsWith(value)
                  );
                case "endsWith":
                  return (
                    typeof cellValue == "string" && cellValue.endsWith(value)
                  );
                default:
                  if (isNullOrEmptyValue(cellValue)) return false;
                  if (typeof cellValue != "number")
                    throw Error(
                      `Ordered filter found a nonnumeric value in “${resolvedCondition.field}” on input record ${filterRowIndex + 1}. Convert types explicitly.`,
                    );
                  return resolvedCondition.operator === "gt"
                    ? cellValue > value
                    : resolvedCondition.operator === "gte"
                      ? cellValue >= value
                      : resolvedCondition.operator === "lt"
                        ? cellValue < value
                        : cellValue <= value;
              }
            },
          );
          return step.join === "and"
            ? conditionResults.every(Boolean)
            : conditionResults.some(Boolean);
        });
      return {
        data: { ...inputTable, rows: rows },
        details: {
          retained: rows.length,
          excluded: inputTable.rows.length - rows.length,
          comparison:
            "Exact typed comparison; text uses case-sensitive Unicode code units.",
        },
      };
    }
    case "sort": {
      const sortKeys = step.fields.map((field) => ({
          ...field,
          i: findFieldIndexOrThrow(inputTable, field.field),
        })),
        typeRank = (sortValue) =>
          sortValue === null
            ? 3
            : typeof sortValue == "number"
              ? 0
              : typeof sortValue == "string"
                ? 1
                : 2,
        rows = inputTable.rows
          .map((row, originalIndex) => ({ r: row, i: originalIndex }))
          .sort((leftEntry, rightEntry) => {
            for (const sortKey of sortKeys) {
              const leftValue = leftEntry.r.values[sortKey.i],
                rightValue = rightEntry.r.values[sortKey.i];
              let comparisonResult = typeRank(leftValue) - typeRank(rightValue);
              if (
                (!comparisonResult &&
                  leftValue !== rightValue &&
                  (comparisonResult = leftValue < rightValue ? -1 : 1),
                comparisonResult)
              )
                return (
                  comparisonResult *
                  (sortKey.direction === "ascending" ? 1 : -1)
                );
            }
            return leftEntry.i - rightEntry.i;
          })
          .map((entry) => entry.r);
      return { data: { ...inputTable, rows: rows } };
    }
    case "unpivot": {
      const unpivotIndices = resolveFieldIndices(step.fields),
        keptIndices = inputTable.fields
          .map((field, fieldIndex) => fieldIndex)
          .filter((candidateIndex) => !unpivotIndices.includes(candidateIndex));
      if (
        inputTable.rows.length * unpivotIndices.length >
        transformationRecipeLimits.rows
      )
        throw Error(
          "Unpivot would exceed 1,000 intermediate rows; narrow the selected fields or source.",
        );
      return {
        data: {
          fields: [
            ...keptIndices.map((keptIndex) => inputTable.fields[keptIndex]),
            step.attribute,
            step.value,
          ],
          rows: inputTable.rows.flatMap((row) =>
            unpivotIndices
              .filter(
                (unpivotableIndex) =>
                  step.keepNulls ||
                  !isNullOrEmptyValue(row.values[unpivotableIndex]),
              )
              .map((unpivotIndex) => ({
                values: [
                  ...keptIndices.map(
                    (keptColumnIndex) => row.values[keptColumnIndex],
                  ),
                  inputTable.fields[unpivotIndex],
                  row.values[unpivotIndex],
                ],
                sources: row.sources,
              })),
          ),
          declaredTypes: { [step.attribute]: "text" },
        },
        details: {
          nullPolicy: step.keepNulls ? "retain" : "exclude",
          selectedFields: step.fields,
        },
      };
    }
    case "group": {
      const { indices, groups } = groupRowsByFields(inputTable, step.keys),
        aggregateSpecs = step.values.map((value) => ({
          ...value,
          i: findFieldIndexOrThrow(inputTable, value.field),
        }));
      return {
        data: {
          fields: [
            ...step.keys,
            ...aggregateSpecs.map((valueSpec) => valueSpec.name),
          ],
          rows: [...groups.values()].map((groupRows) => ({
            values: [
              ...indices.map((indice) => groupRows[0].values[indice]),
              ...aggregateSpecs.map((aggregateSpec) =>
                aggregateFieldValues(
                  groupRows,
                  aggregateSpec.i,
                  aggregateSpec.operation,
                ),
              ),
            ],
            sources: uniqueSortedSources(groupRows),
          })),
          declaredTypes: {},
        },
      };
    }
    case "pivot": {
      if (
        step.keys.includes(step.column) ||
        step.keys.includes(step.value) ||
        step.column === step.value
      )
        throw Error("Pivot keys, attribute and value fields must be distinct.");
      const fieldIndexOrThrow = findFieldIndexOrThrow(inputTable, step.column),
        valueIndex = findFieldIndexOrThrow(inputTable, step.value),
        { indices, groups } = groupRowsByFields(inputTable, step.keys),
        pivotColumnValues = [
          ...new Map(
            inputTable.rows.map((row) => [
              serializeTypedKeyValues([row.values[fieldIndexOrThrow]]),
              row.values[fieldIndexOrThrow],
            ]),
          ).values(),
        ],
        pivotHeaders = pivotColumnValues.map(
          (pivotValue) =>
            `${pivotValue === null ? "blank" : typeof pivotValue}:${pivotValue === null ? "(blank)" : String(pivotValue)}`,
        );
      if (step.keys.length + pivotHeaders.length > 50)
        throw Error("Pivot exceeds 50 output columns.");
      const rows = [...groups.values()].map((pivotGroupRows) => ({
        values: [
          ...indices.map((indice) => pivotGroupRows[0].values[indice]),
          ...pivotColumnValues.map((columnValue) =>
            aggregateFieldValues(
              pivotGroupRows.filter(
                (groupRow) =>
                  groupRow.values[fieldIndexOrThrow] === columnValue,
              ),
              valueIndex,
              step.operation,
            ),
          ),
        ],
        sources: uniqueSortedSources(pivotGroupRows),
      }));
      if (!pivotColumnValues.length && step.keys.length === 0)
        throw Error(
          "An empty pivot has no columns. Select at least one grouping key.",
        );
      return {
        data: {
          fields: [...step.keys, ...pivotHeaders],
          rows: rows,
          declaredTypes: {},
        },
        details: {
          columns: pivotColumnValues.map(
            (pivotColumnValue, pivotColumnIndex) => ({
              header: pivotHeaders[pivotColumnIndex],
              value: pivotColumnValue,
            }),
          ),
          aggregation: step.operation,
        },
      };
    }
    case "append": {
      const appendSourceTable = loadSourceTable(step.source);
      if (
        inputTable.fields.length !== appendSourceTable.fields.length ||
        inputTable.fields.some(
          (field) => !appendSourceTable.fields.includes(field),
        )
      )
        throw Error(
          "Append requires the same exact header set; column order may differ.",
        );
      const appendColumnIndices = inputTable.fields.map((field) =>
        findFieldIndexOrThrow(appendSourceTable, field),
      );
      return {
        data: {
          fields: inputTable.fields,
          rows: [
            ...inputTable.rows,
            ...appendSourceTable.rows.map((row) => ({
              ...row,
              values: appendColumnIndices.map(
                (sourceColumnIndex) => row.values[sourceColumnIndex],
              ),
            })),
          ],
          declaredTypes: {},
        },
        details: {
          leftRows: inputTable.rows.length,
          rightRows: appendSourceTable.rows.length,
          outputRows: inputTable.rows.length + appendSourceTable.rows.length,
        },
      };
    }
    case "merge": {
      const mergeSourceTable = loadSourceTable(step.source),
        leftKeyIndices = step.keys.map((key) =>
          findFieldIndexOrThrow(inputTable, key.left),
        ),
        rightKeyIndices = step.keys.map((key) =>
          findFieldIndexOrThrow(mergeSourceTable, key.right),
        );
      if (
        new Set(leftKeyIndices).size !== leftKeyIndices.length ||
        new Set(rightKeyIndices).size !== rightKeyIndices.length
      )
        throw Error("Merge keys cannot repeat a field.");
      const toMatchRecords = (matchTable, tableLabel) =>
          matchTable.rows.map((row, matchRowIndex) => ({
            row: matchRowIndex,
            locator: `${tableLabel} input record ${matchRowIndex + 1}`,
            values: row.values,
          })),
        records = toMatchRecords(inputTable, "Left"),
        rightRecords = toMatchRecords(mergeSourceTable, "Right"),
        keyMatches = matchKeyPopulations(
          records,
          rightRecords,
          leftKeyIndices,
          rightKeyIndices,
          step.normalization,
          step.fuzzy,
        ),
        rows = [],
        pushMergedRow = (leftRow, rightRow) => {
          if (rows.length >= 1e3)
            throw Error("Merge exceeds 1,000 intermediate records.");
          rows.push({
            values: [
              ...(leftRow?.values ?? inputTable.fields.map(() => null)),
              ...(rightRow?.values ?? mergeSourceTable.fields.map(() => null)),
            ],
            sources: uniqueSortedSources(
              [leftRow, rightRow].filter((candidateRow) => !!candidateRow),
            ),
          });
        };
      for (const record of records) {
        const leftMatches = keyMatches.byLeft.get(record.row);
        if (leftMatches) {
          if (!["leftAnti", "rightAnti"].includes(step.mode))
            for (const leftMatch of leftMatches)
              pushMergedRow(
                inputTable.rows[record.row],
                mergeSourceTable.rows[leftMatch.record.row],
              );
        } else
          ["left", "full", "leftAnti"].includes(step.mode) &&
            pushMergedRow(inputTable.rows[record.row], undefined);
      }
      if (["right", "full", "rightAnti"].includes(step.mode))
        for (const unmatchedRecord of keyMatches.unmatchedRight)
          pushMergedRow(undefined, mergeSourceTable.rows[unmatchedRecord.row]);
      return {
        data: {
          fields: [
            ...inputTable.fields,
            ...mergeSourceTable.fields.map((field) => step.prefix + field),
          ],
          rows: rows,
          declaredTypes: {},
        },
        details: {
          leftRows: inputTable.rows.length,
          rightRows: mergeSourceTable.rows.length,
          matchedLeft: keyMatches.matchedLeftRows,
          matchedRight: keyMatches.matchedRightRows,
          matchPairs: keyMatches.matchPairs,
          unmatchedLeft: keyMatches.unmatchedLeft,
          unmatchedRight: keyMatches.unmatchedRight,
          duplicateLeft: keyMatches.duplicateLeftKeys,
          duplicateRight: keyMatches.duplicateRightKeys,
          manyToMany: keyMatches.manyToManyKeys,
          extraPairs: keyMatches.expandedMatchRows,
          matching: keyMatches.diagnostics,
        },
      };
    }
  }
}

export {
  executeRecipeStep,
  commitApprovedTransformation,
  discoverArrayFormulaAnchors,
  resolveRecipeSourceTable,
  verifyTransformationPreview,
  applyTextSplitPreview,
  selectExceptionFreeRows,
  buildTransformationPreview,
  validateRecipeOutputTable,
  buildTextSplitResult,
  describeInferredTextRule,
  profileFieldTypes,
  buildFieldControlTotals,
};
export { coerceCellValueToText };
