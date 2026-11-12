import { isPlainObjectValue } from "./01-cell-styles-and-themes.js";
import { isFiniteNumber } from "./04-pivot-tables-and-pictures.js";
import {
  workbookLimits,
  maxSelectionAreas,
} from "../02-text-and-financial-functions.js";
import {
  hasLinkedOrBoundCells,
  isoDateTimeSchemaFactory,
} from "../01-schemas-date-functions-and-validation.js";
import {
  string,
  number,
  boolean,
  array,
  strictObject,
  literal,
  enum as _enum,
  unknown,
} from "zod";
import {
  validateSplitPaneState,
  normalizeHiddenIntervals,
  conditionalRuleKinds,
  readHiddenAxisIntervals,
  isNonArrayObject,
  hiddenKeyForAxis,
  columnWidthAt,
  sortConditionalRulesByPriority,
  mapIndexAfterMove,
  reviewThreadLimits,
  normalizeReviewTimestamp,
  axisCountOf,
  addOutlineGroup,
  validateOutlineGroups,
  remapOutlineGroupsForEdit,
  isErrorValueText,
  assertBoundedString,
  assertPlainReviewRecord,
  unsafeControlCharPattern,
  assertWorksheetAxis,
  validationLimits,
  validateCellBorder,
  cn,
  resolveReviewTimestamp,
  rowHeightAt,
  throwWithMessage,
  assertSheetUnprotected,
  normalizeBoundedText,
  assertSixDigitHexColor,
  validateSheetDefaultSizes,
  isCellInsideRange,
  cellBorderEdges,
  createAxisVisibilityIndex,
  failWithMessage,
  vx,
  normalizeReviewAuthor,
  worksheetCapacityLimits,
  assertBoundedValidationText,
} from "./01-cell-styles-and-themes.js";
import {
  sumDecimalPartsToNumber,
  exactAverageOfNumbers,
  exactProductOfNumbers,
} from "./02-formula-function-evaluators.js";
import {
  throwXlsxDefinedNameError,
  spreadsheetMlNamespace,
  stripFormulaStringLiterals,
  validateXmlSafeName,
  parseDefinedNameXmlAttributes,
  validateDefinedName,
  parseDefinedNamesXml,
  nP,
  validateNameFormulaText,
  reservedDefinedNames,
} from "./03-defined-names.js";
import {
  containsSubtotalCall,
  inspectPictureDataUrl,
  containsSubtotalOrAggregateCall,
  referenceSensitiveFunctionNames,
  notReferenceFunctionMarker,
  isFormulaCellValue,
} from "./04-pivot-tables-and-pictures.js";
import {
  B2,
  requireWhatIfName,
  HO,
  transformCellValueSchema,
  whatIfLimits,
  expandWhatIfScenarios,
  referenceImpactKinds,
  cellAddressKey,
  validateWhatIfFields,
  isBoundedPlainText,
  validateWhatIfValues,
  requireWhatIfList,
  columnIndexToLetters,
  printRenderLimits,
  shiftFormulaReferences,
} from "./06-data-model-and-analysis.js";

function resolveFilteredRowContext(workbook, options = {}) {
  if (
    !isPlainObjectValue(options) ||
    Object.keys(options).some(
      (optionKey) => optionKey !== "filteredRowsBySheet",
    )
  )
    throw Error("Calculation options support only filteredRowsBySheet.");
  const filteredRowsBySheet = options.filteredRowsBySheet;
  if (
    filteredRowsBySheet !== undefined &&
    !isPlainObjectValue(filteredRowsBySheet)
  )
    throw Error(
      "Filtered rows must be a worksheet-ID map of inclusive row intervals.",
    );
  const sheetsById = new Map(workbook.sheets.map((sheet) => [sheet.id, sheet])),
    sheetsByName = new Map(
      workbook.sheets.map((sheet) => [sheet.name.toLowerCase(), sheet]),
    ),
    filterRowIndexBySheetId = new Map();
  let intervalCount = 0;
  for (const [sheetId, intervals] of Object.entries(
    filteredRowsBySheet ?? {},
  )) {
    const targetSheet = sheetsById.get(sheetId);
    if (!targetSheet)
      throw Error("Filter context refers to an unknown worksheet ID.");
    if (!Array.isArray(intervals) || (intervalCount += intervals.length) > 1e4)
      throw Error(
        "Filter context supports at most 10,000 intervals across the workbook.",
      );
    filterRowIndexBySheetId.set(
      sheetId,
      createAxisVisibilityIndex(
        targetSheet.rowCount,
        normalizeHiddenIntervals(intervals, targetSheet.rowCount),
      ),
    );
  }
  const hiddenRowIndexBySheetId = new Map();
  for (const sheet of workbook.sheets)
    hiddenRowIndexBySheetId.set(
      sheet.id,
      createAxisVisibilityIndex(sheet.rowCount, cn(sheet, "row")),
    );
  return (functionName, args, context) => {
    if (!referenceSensitiveFunctionNames.includes(functionName))
      return notReferenceFunctionMarker;
    const { issue, isIssue } = context,
      resolveReferenceRange = (argumentValue) => {
        if (isIssue(argumentValue)) throw argumentValue;
        if (
          !isPlainObjectValue(argumentValue) ||
          !isPlainObjectValue(argumentValue.ref)
        )
          throw issue("#VALUE!");
        const ref = argumentValue.ref;
        if (ref.sheet !== undefined && typeof ref.sheet != "string")
          throw issue("#REF!");
        const sheet = sheetsByName.get(
          String(ref.sheet ?? context.currentSheetName ?? "").toLowerCase(),
        );
        if (!sheet) throw issue("#REF!");
        const fromRef = isPlainObjectValue(ref.from) ? ref.from : ref,
          toRef = isPlainObjectValue(ref.to) ? ref.to : ref,
          bounds = [fromRef.row, fromRef.col, toRef.row, toRef.col];
        if (bounds.some((coordinate) => !Number.isInteger(coordinate)))
          throw issue("#REF!");
        const [firstRow, firstCol, lastRow, lastCol] = bounds;
        if (
          firstRow < 1 ||
          firstCol < 1 ||
          lastRow < firstRow ||
          lastCol < firstCol ||
          lastRow > sheet.rowCount ||
          lastCol > sheet.colCount
        )
          throw issue("#REF!");
        return {
          sheet: sheet,
          firstRow: firstRow,
          lastRow: lastRow,
          firstCol: firstCol,
          lastCol: lastCol,
        };
      };
    if (functionName === "FORMULATEXT" || functionName === "ISFORMULA") {
      if (args.length !== 1) return issue("#VALUE!");
      const targetRange = resolveReferenceRange(args[0]),
        readCellAt = (rowNumber, columnNumber) => {
          const cellAddress = `${columnIndexToLetters(columnNumber - 1)}${rowNumber}`;
          return (
            context.readMetadata?.(targetRange.sheet.id, cellAddress),
            targetRange.sheet.cells[cellAddress]
          );
        };
      if (functionName === "FORMULATEXT") {
        const cell = readCellAt(targetRange.firstRow, targetRange.firstCol);
        return isFormulaCellValue(cell) && cell.raw.length <= 8192
          ? cell.raw
          : issue("#N/A");
      }
      const cellCount =
        (targetRange.lastRow - targetRange.firstRow + 1) *
        (targetRange.lastCol - targetRange.firstCol + 1);
      return cellCount > 1e4
        ? issue("#LIMIT!")
        : cellCount === 1
          ? isFormulaCellValue(
              readCellAt(targetRange.firstRow, targetRange.firstCol),
            )
          : Array.from(
              { length: targetRange.lastRow - targetRange.firstRow + 1 },
              (K, rowOffset) =>
                Array.from(
                  { length: targetRange.lastCol - targetRange.firstCol + 1 },
                  (ae, columnOffset) =>
                    isFormulaCellValue(
                      readCellAt(
                        targetRange.firstRow + rowOffset,
                        targetRange.firstCol + columnOffset,
                      ),
                    ),
                ),
            );
    }
    const isAggregate = functionName === "AGGREGATE";
    if (args.length < (isAggregate ? 3 : 2) || args.length > 255)
      return issue("#VALUE!");
    const readNumberArgument = (argument) => {
        const retrievedValue = context.retrieve(argument);
        if (isIssue(retrievedValue)) throw retrievedValue;
        if (
          Array.isArray(retrievedValue) ||
          !["number", "string", "boolean"].includes(typeof retrievedValue) ||
          (typeof retrievedValue == "string" && retrievedValue.trim() === "") ||
          !Number.isFinite(Number(retrievedValue))
        )
          throw issue("#VALUE!");
        return Number(retrievedValue);
      },
      rawFunctionCode = readNumberArgument(args[0]);
    if (isAggregate && !Number.isInteger(rawFunctionCode))
      return issue("#VALUE!");
    const truncatedFunctionCode = Math.trunc(rawFunctionCode),
      functionCode =
        !isAggregate && truncatedFunctionCode > 100
          ? truncatedFunctionCode - 100
          : truncatedFunctionCode;
    if (functionCode < 1 || functionCode > (isAggregate ? 19 : 11))
      return issue("#VALUE!");
    const aggregateOption = isAggregate
      ? args[1] == null
        ? 0
        : readNumberArgument(args[1])
      : 0;
    if (
      !Number.isInteger(aggregateOption) ||
      aggregateOption < 0 ||
      aggregateOption > 7
    )
      return issue("#VALUE!");
    const shouldExcludeHiddenRows = isAggregate
        ? aggregateOption % 2 === 1
        : truncatedFunctionCode > 100,
      shouldIgnoreErrors =
        isAggregate && [2, 3, 6, 7].includes(aggregateOption),
      shouldSkipNestedSubtotals = !isAggregate || aggregateOption < 4,
      hasKArgument = isAggregate && functionCode >= 14;
    if (hasKArgument && (args.length !== 4 || args[3] == null))
      return issue("#VALUE!");
    const kValue = hasKArgument ? readNumberArgument(args[3]) : 0,
      rangeArgs = hasKArgument ? [args[2]] : args.slice(isAggregate ? 2 : 1),
      isArrayForm =
        hasKArgument &&
        !((possibleRef) =>
          isPlainObjectValue(possibleRef) &&
          isPlainObjectValue(possibleRef.ref))(rangeArgs[0]),
      referenceRanges = isArrayForm ? [] : rangeArgs.map(resolveReferenceRange);
    let totalCellCount = 0;
    for (const sizedRange of referenceRanges) {
      const rangeCellCount =
        (sizedRange.lastRow - sizedRange.firstRow + 1) *
        (sizedRange.lastCol - sizedRange.firstCol + 1);
      if (
        ((totalCellCount += rangeCellCount),
        rangeCellCount > 2e4 || totalCellCount > 1e5)
      )
        return issue("#LIMIT!");
    }
    for (const cycleRange of referenceRanges)
      for (
        let firstRow = cycleRange.firstRow;
        firstRow <= cycleRange.lastRow;
        firstRow++
      )
        for (
          let firstCol = cycleRange.firstCol;
          firstCol <= cycleRange.lastCol;
          firstCol++
        )
          if (
            context.isActive(
              cycleRange.sheet.id,
              `${columnIndexToLetters(firstCol - 1)}${firstRow}`,
            )
          )
            return issue("#CYCLE!");
    let numberCount = 0,
      nonBlankCount = 0,
      minimum = 1 / 0,
      maximum = -1 / 0,
      runningMean = 0,
      squaredDeviationSum = 0;
    const collectedValues = [],
      accumulateValue = (value) => {
        if (isIssue(value)) {
          if (["#LIMIT!", "#CYCLE!"].includes(value.error)) return value;
          if (shouldIgnoreErrors) return;
          if (functionCode !== 2 && functionCode !== 3) return value;
        }
        if ((value != null && nonBlankCount++, typeof value != "number"))
          return;
        if (!Number.isFinite(value)) return issue("#NUM!");
        (numberCount++,
          (minimum = Math.min(minimum, value)),
          (maximum = Math.max(maximum, value)));
        const delta = value - runningMean;
        ((runningMean += delta / numberCount),
          (squaredDeviationSum += delta * (value - runningMean)),
          (functionCode === 1 ||
            functionCode === 6 ||
            functionCode === 9 ||
            functionCode >= 12) &&
            collectedValues.push(value));
      };
    if (isArrayForm) {
      const arrayArgument = context.retrieve(rangeArgs[0]),
        rows = Array.isArray(arrayArgument) ? arrayArgument : [[arrayArgument]];
      if (!rows.length || rows.length > 1e4) return issue("#LIMIT!");
      const grid = rows.every(Array.isArray) ? rows : [rows],
        columnCount = grid[0]?.length ?? 0;
      if (
        !columnCount ||
        grid.some(
          (gridRow) =>
            gridRow.length !== columnCount || gridRow.some(Array.isArray),
        )
      )
        return issue("#VALUE!");
      if (grid.length * columnCount > 1e4) return issue("#LIMIT!");
      for (const arrayRow of grid)
        for (const arrayCell of arrayRow) {
          if (
            arrayCell != null &&
            !isIssue(arrayCell) &&
            !["number", "string", "boolean"].includes(typeof arrayCell)
          )
            return issue("#VALUE!");
          const accumulationError = accumulateValue(arrayCell);
          if (accumulationError) return accumulationError;
        }
    }
    for (const scannedRange of referenceRanges)
      for (
        let firstRow = scannedRange.firstRow;
        firstRow <= scannedRange.lastRow;
        firstRow++
      )
        if (!(
          filterRowIndexBySheetId
            .get(scannedRange.sheet.id)
            ?.isHidden(firstRow - 1) ||
          (shouldExcludeHiddenRows &&
            hiddenRowIndexBySheetId
              .get(scannedRange.sheet.id)
              .isHidden(firstRow - 1))
        ))
          for (
            let firstCol = scannedRange.firstCol;
            firstCol <= scannedRange.lastCol;
            firstCol++
          ) {
            const cellAddressText = `${columnIndexToLetters(firstCol - 1)}${firstRow}`;
            context.readMetadata?.(scannedRange.sheet.id, cellAddressText);
            const storedCell = scannedRange.sheet.cells[cellAddressText];
            if (
              shouldSkipNestedSubtotals &&
              isFormulaCellValue(storedCell) &&
              (isAggregate
                ? containsSubtotalOrAggregateCall
                : containsSubtotalCall)(
                context.expandedFormula(storedCell, scannedRange.sheet.id),
              )
            )
              continue;
            const cellError = accumulateValue(
              context.readCell(scannedRange.sheet.id, cellAddressText),
            );
            if (cellError) return cellError;
          }
    if (functionCode >= 12) {
      if (
        (collectedValues.sort(
          (leftValue, rightValue) => leftValue - rightValue,
        ),
        functionCode === 13)
      ) {
        let maxRunLength = 1,
          modeValue = 0,
          currentRunLength = 0,
          previousValue;
        for (const sortedValue of collectedValues)
          ((currentRunLength =
            sortedValue === previousValue ? currentRunLength + 1 : 1),
            currentRunLength > maxRunLength &&
              ((maxRunLength = currentRunLength), (modeValue = sortedValue)),
            (previousValue = sortedValue));
        return maxRunLength > 1 ? modeValue : issue("#N/A");
      }
      if (!collectedValues.length) return issue("#NUM!");
      if (functionCode === 12) {
        const middleIndex = Math.floor(collectedValues.length / 2);
        return collectedValues.length % 2
          ? collectedValues[middleIndex]
          : collectedValues[middleIndex - 1] / 2 +
              collectedValues[middleIndex] / 2;
      }
      if (functionCode === 14 || functionCode === 15) {
        const kIndex = kValue;
        return Number.isInteger(kIndex) &&
          kIndex >= 1 &&
          kIndex <= collectedValues.length
          ? collectedValues[
              functionCode === 14 ? collectedValues.length - kIndex : kIndex - 1
            ]
          : issue("#NUM!");
      }
      const isExclusive = functionCode >= 18,
        percentileFraction =
          functionCode === 17 || functionCode === 19
            ? Math.trunc(kValue) / 4
            : kValue;
      if (
        percentileFraction < 0 ||
        percentileFraction > 1 ||
        (isExclusive && (percentileFraction === 0 || percentileFraction === 1))
      )
        return issue("#NUM!");
      const fractionalIndex = isExclusive
        ? percentileFraction * (collectedValues.length + 1) - 1
        : percentileFraction * (collectedValues.length - 1);
      if (fractionalIndex < 0 || fractionalIndex > collectedValues.length - 1)
        return issue("#NUM!");
      const lowerIndex = Math.floor(fractionalIndex),
        upperIndex = Math.ceil(fractionalIndex),
        interpolationWeight = fractionalIndex - lowerIndex;
      return (
        collectedValues[lowerIndex] * (1 - interpolationWeight) +
        collectedValues[upperIndex] * interpolationWeight
      );
    }
    if (functionCode === 2) return numberCount;
    if (functionCode === 3) return nonBlankCount;
    if (functionCode === 9) return sumDecimalPartsToNumber(collectedValues);
    if (functionCode === 1)
      return numberCount
        ? exactAverageOfNumbers(collectedValues)
        : issue("#DIV/0!");
    if (functionCode === 4) return numberCount ? maximum : 0;
    if (functionCode === 5) return numberCount ? minimum : 0;
    if (functionCode === 6)
      return numberCount ? exactProductOfNumbers(collectedValues) : 0;
    const isSampleStatistic = functionCode === 7 || functionCode === 10;
    if (numberCount < (isSampleStatistic ? 2 : 1)) return issue("#DIV/0!");
    const variance = Math.max(
      0,
      squaredDeviationSum / (numberCount - (isSampleStatistic ? 1 : 0)),
    );
    return functionCode === 7 || functionCode === 8
      ? Math.sqrt(variance)
      : variance;
  };
}
function formatAxisIntervalRef(interval, axis) {
  return axis === "row"
    ? `${interval.start + 1}:${interval.end + 1}`
    : `${columnIndexToLetters(interval.start)}:${columnIndexToLetters(interval.end)}`;
}
function _t(address) {
  const addressMatch = /^\$?([A-Za-z]{1,3})\$?([1-9]\d{0,6})$/.exec(address);
  if (!addressMatch)
    return throwWithMessage(`Invalid cell address: ${address}`);
  let columnNumber = 0;
  for (const letter of addressMatch[1].toUpperCase())
    columnNumber = columnNumber * 26 + letter.charCodeAt(0) - 64;
  const row = Number(addressMatch[2]) - 1;
  return row >= workbookLimits.maxRows || columnNumber > workbookLimits.maxCols
    ? throwWithMessage(`Cell ${address} exceeds the supported workbook bounds.`)
    : { row: row, col: columnNumber - 1 };
}
function buildBorderPresetStyles(
  address,
  sampleGridAddresses,
  preset,
  line = { style: "thin", color: "#000000" },
) {
  const anchorCell = _t(address),
    validatedBorder = validateCellBorder(line),
    presetBorders = {};
  ["all", "outline", "inside", "clear"].includes(preset) ||
    failWithMessage("Choose a supported border preset.");
  for (const cellBorderEdge of cellBorderEdges) {
    const neighborRow =
        anchorCell.row +
        (cellBorderEdge === "top" ? -1 : cellBorderEdge === "bottom" ? 1 : 0),
      neighborColumn =
        anchorCell.col +
        (cellBorderEdge === "left" ? -1 : cellBorderEdge === "right" ? 1 : 0),
      hasNeighbor =
        neighborRow >= 0 &&
        neighborRow < 1048576 &&
        neighborColumn >= 0 &&
        neighborColumn < 16384 &&
        sampleGridAddresses.has(
          `${columnIndexToLetters(neighborColumn)}${neighborRow + 1}`,
        );
    preset === "clear"
      ? (presetBorders[cellBorderEdge] = { style: "none" })
      : (preset === "all" ||
          (preset === "outline" ? !hasNeighbor : hasNeighbor)) &&
        (presetBorders[cellBorderEdge] = { ...validatedBorder });
  }
  return presetBorders;
}
function normalizeCellAddress(address) {
  if (typeof address != "string")
    return throwWithMessage("Choose a valid worksheet cell address.");
  const cellPosition = _t(address.trim());
  return `${columnIndexToLetters(cellPosition.col)}${cellPosition.row + 1}`;
}
function validateHyperlinkRecord(hyperlink) {
  const hyperlinkRecord = assertPlainReviewRecord(hyperlink, "Hyperlink"),
    normalizedLabel =
      hyperlinkRecord.label === undefined || hyperlinkRecord.label === ""
        ? undefined
        : normalizeBoundedText(
            hyperlinkRecord.label,
            "Link display text",
            reviewThreadLimits.labelLength,
          );
  if (hyperlinkRecord.kind === "internal")
    return {
      kind: "internal",
      sheetId: normalizeBoundedText(
        hyperlinkRecord.sheetId,
        "Link worksheet ID",
        100,
      ),
      address: normalizeCellAddress(hyperlinkRecord.address),
      ...(normalizedLabel ? { label: normalizedLabel } : {}),
    };
  if (hyperlinkRecord.kind !== "external")
    return throwWithMessage("Choose a web, email or worksheet link.");
  const boundedText = normalizeBoundedText(
    hyperlinkRecord.url,
    "Link address",
    reviewThreadLimits.urlLength,
  );
  let decodedUrl = boundedText;
  for (let decodePass = 0; decodePass < 4; decodePass++) {
    (unsafeControlCharPattern.test(decodedUrl) || /\\/.test(decodedUrl)) &&
      throwWithMessage(
        "Link addresses cannot contain control characters or backslashes.",
      );
    let nextUrl = decodedUrl;
    try {
      nextUrl = decodeURIComponent(decodedUrl);
    } catch {
      throwWithMessage("The link contains invalid encoded characters.");
    }
    if (nextUrl === decodedUrl) break;
    decodedUrl = nextUrl;
  }
  (unsafeControlCharPattern.test(decodedUrl) ||
    /%(?:0[\da-f]|1[\da-f]|7f|8[\da-f]|9[\da-f])/i.test(decodedUrl)) &&
    throwWithMessage(
      "Link addresses cannot contain encoded control characters.",
    );
  let parsedUrl;
  try {
    parsedUrl = new URL(boundedText);
  } catch {
    return throwWithMessage(
      "Enter a complete https://, http:// or mailto: address.",
    );
  }
  return (
    ["https:", "http:", "mailto:"].includes(parsedUrl.protocol) ||
      throwWithMessage(
        "Only https://, http:// and mailto: links are supported.",
      ),
    (parsedUrl.username || parsedUrl.password) &&
      throwWithMessage("Links cannot contain embedded credentials."),
    parsedUrl.protocol === "mailto:"
      ? (parsedUrl.search ||
          parsedUrl.hash ||
          parsedUrl.host ||
          !/^[^\s@/?#:,]+@[^\s@/?#:,]+\.[^\s@/?#:,]+$/.test(
            decodeURIComponent(parsedUrl.pathname),
          )) &&
        throwWithMessage(
          "Use mailto: with one email address and no additional headers.",
        )
      : parsedUrl.hostname || throwWithMessage("Web links require a hostname."),
    parsedUrl.href.length > reviewThreadLimits.urlLength &&
      throwWithMessage("The normalized link is too long."),
    {
      kind: "external",
      url: parsedUrl.href,
      ...(normalizedLabel ? { label: normalizedLabel } : {}),
    }
  );
}
function normalizeReviewMetadata(metadata, workbook) {
  const plainMetadata = assertPlainReviewRecord(
    metadata,
    "Workbook review metadata",
  );
  if (plainMetadata.reviewThreads !== undefined) {
    (!Array.isArray(plainMetadata.reviewThreads) ||
      plainMetadata.reviewThreads.length > reviewThreadLimits.threads) &&
      throwWithMessage(
        `A workbook can contain at most ${reviewThreadLimits.threads} review threads.`,
      );
    const threadIds = new Set(),
      messageIds = new Set();
    let totalMessages = 0;
    const threads = plainMetadata.reviewThreads.map((reviewThread) => {
      const threadRecord = assertPlainReviewRecord(
          reviewThread,
          "Review thread",
        ),
        anchorRecord = assertPlainReviewRecord(
          threadRecord.anchor,
          "Review target",
        ),
        boundedText = normalizeBoundedText(
          threadRecord.id,
          "Review thread ID",
          100,
        );
      (threadIds.has(boundedText) &&
        throwWithMessage("Review thread IDs must be unique."),
        threadIds.add(boundedText));
      const reviewTimestamp = normalizeReviewTimestamp(threadRecord.createdAt),
        updatedTimestamp = normalizeReviewTimestamp(threadRecord.updatedAt);
      updatedTimestamp < reviewTimestamp &&
        throwWithMessage("Review update time cannot precede creation.");
      const anchor = {
        sheetId: normalizeBoundedText(
          anchorRecord.sheetId,
          "Review worksheet ID",
          100,
        ),
        sheetName: normalizeBoundedText(
          anchorRecord.sheetName,
          "Review worksheet name",
          31,
        ),
        address: normalizeCellAddress(anchorRecord.address),
      };
      (!Array.isArray(threadRecord.messages) ||
        !threadRecord.messages.length ||
        threadRecord.messages.length > reviewThreadLimits.messagesPerThread ||
        (totalMessages += threadRecord.messages.length) >
          reviewThreadLimits.totalMessages) &&
        throwWithMessage(
          "Review discussions exceed the per-thread or workbook message limit.",
        );
      let previousCreatedAt = reviewTimestamp;
      const messages = threadRecord.messages.map((message, messageIndex) => {
          const messageRecord = assertPlainReviewRecord(
              message,
              "Review message",
            ),
            boundedText = normalizeBoundedText(
              messageRecord.id,
              "Review message ID",
              100,
            );
          (messageIds.has(boundedText) &&
            throwWithMessage(
              "Review message IDs must be unique within the workbook.",
            ),
            messageIds.add(boundedText));
          const messageCreatedAt = normalizeReviewTimestamp(
            messageRecord.createdAt,
          );
          return (
            (messageCreatedAt < previousCreatedAt ||
              messageCreatedAt > updatedTimestamp ||
              (messageIndex === 0 && messageCreatedAt !== reviewTimestamp)) &&
              throwWithMessage(
                "Review message times must follow their discussion in order.",
              ),
            (previousCreatedAt = messageCreatedAt),
            {
              id: boundedText,
              author: normalizeReviewAuthor(
                normalizeBoundedText(
                  messageRecord.author,
                  "Reviewer name",
                  100,
                ),
              ),
              text: normalizeBoundedText(
                messageRecord.text,
                "Comment",
                reviewThreadLimits.textLength,
                true,
              ),
              createdAt: messageCreatedAt,
            }
          );
        }),
        thread = {
          id: boundedText,
          anchor: anchor,
          messages: messages,
          createdAt: reviewTimestamp,
          updatedAt: updatedTimestamp,
        };
      if (threadRecord.resolved !== undefined) {
        const resolutionRecord = assertPlainReviewRecord(
            threadRecord.resolved,
            "Review resolution",
          ),
          reviewTimestamp = normalizeReviewTimestamp(
            resolutionRecord.createdAt,
          );
        ((reviewTimestamp < messages.at(-1).createdAt ||
          reviewTimestamp > updatedTimestamp) &&
          throwWithMessage("Invalid review resolution time."),
          (thread.resolved = {
            author: normalizeReviewAuthor(
              normalizeBoundedText(
                resolutionRecord.author,
                "Reviewer name",
                100,
              ),
            ),
            createdAt: reviewTimestamp,
          }));
      }
      if (threadRecord.orphaned !== undefined) {
        const orphanedRecord = assertPlainReviewRecord(
            threadRecord.orphaned,
            "Orphaned review target",
          ),
          orphanedTimestamp = normalizeReviewTimestamp(
            orphanedRecord.createdAt,
          );
        ((!["worksheet-deleted", "cell-deleted"].includes(
          String(orphanedRecord.reason),
        ) ||
          orphanedTimestamp < reviewTimestamp ||
          orphanedTimestamp > updatedTimestamp) &&
          throwWithMessage("Invalid orphaned review target."),
          (thread.orphaned = {
            reason: orphanedRecord.reason,
            createdAt: orphanedTimestamp,
          }));
      }
      return (
        !thread.orphaned &&
          !workbook.sheets.some((sheet) => sheet.id === anchor.sheetId) &&
          throwWithMessage(
            "A review thread with a missing worksheet must be retained as orphaned.",
          ),
        thread
      );
    });
    threads.length
      ? (workbook.reviewThreads = threads)
      : delete workbook.reviewThreads;
  }
  for (const sheet of workbook.sheets)
    for (const cell of hasLinkedOrBoundCells(sheet.cells)
      ? Object.values(sheet.cells)
      : []) {
      if (cell.hyperlink === undefined) continue;
      const hyperlink = validateHyperlinkRecord(cell.hyperlink);
      (hyperlink.kind === "internal" &&
        !workbook.sheets.some((sheet) => sheet.id === hyperlink.sheetId) &&
        throwWithMessage(
          "An internal hyperlink must target an existing worksheet.",
        ),
        (cell.hyperlink = hyperlink));
    }
}
function attachReviewThreadsToWorkbook(sourceWorkbook, reviewThreads) {
  const workbook = { ...sourceWorkbook, reviewThreads: reviewThreads },
    clonedWorkbook = {
      ...workbook,
      sheets: workbook.sheets.map((sheet) => ({
        ...sheet,
        cells: Object.fromEntries(
          Object.entries(sheet.cells).map(([address, cell]) => [
            address,
            { ...cell },
          ]),
        ),
      })),
    };
  return (normalizeReviewMetadata(workbook, clonedWorkbook), clonedWorkbook);
}
function shiftReviewTargetsOnAxisEdit(workbook, sheetId, change, timestamp) {
  const axisLimit =
    change.axis === "row" ? workbookLimits.maxRows : workbookLimits.maxCols;
  (!["row", "column"].includes(change.axis) ||
    !Number.isInteger(change.index) ||
    change.index < 0 ||
    !Number.isInteger(change.count) ||
    change.count < 1 ||
    change.index + change.count > axisLimit) &&
    throwWithMessage("Choose a valid row or column change for review targets.");
  const shiftAddress = (sourceAddress) => {
    const parsedAddress = _t(sourceAddress),
      axisIndex = change.axis === "row" ? parsedAddress.row : parsedAddress.col;
    if (change.moveBefore === undefined && axisIndex < change.index)
      return sourceAddress;
    if (
      change.moveBefore === undefined &&
      change.remove &&
      axisIndex < change.index + change.count
    )
      return null;
    const shiftedIndex =
      change.moveBefore !== undefined
        ? mapIndexAfterMove(axisIndex, change)
        : axisIndex + (change.remove ? -change.count : change.count);
    return (
      shiftedIndex >= axisLimit &&
        throwWithMessage(
          "Insertion would discard a review target or hyperlink at the worksheet boundary.",
        ),
      change.axis === "row"
        ? `${columnIndexToLetters(parsedAddress.col)}${shiftedIndex + 1}`
        : `${columnIndexToLetters(shiftedIndex)}${parsedAddress.row + 1}`
    );
  };
  let hasChanges = false;
  const reviewThreads = workbook.reviewThreads?.map((reviewThread) => {
      if (reviewThread.orphaned || reviewThread.anchor.sheetId !== sheetId)
        return reviewThread;
      const address = shiftAddress(reviewThread.anchor.address);
      if (address === reviewThread.anchor.address) return reviewThread;
      if (((hasChanges = true), address))
        return {
          ...reviewThread,
          anchor: { ...reviewThread.anchor, address: address },
        };
      const reviewTimestamp = resolveReviewTimestamp(
        { timestamp: timestamp },
        reviewThread.updatedAt,
      );
      return {
        ...reviewThread,
        updatedAt: reviewTimestamp,
        orphaned: { reason: "cell-deleted", createdAt: reviewTimestamp },
      };
    }),
    sheets = workbook.sheets.map((sheet) => {
      let cells;
      for (const [cellAddress, cell] of Object.entries(sheet.cells)) {
        const hyperlink = cell.hyperlink;
        if (hyperlink?.kind !== "internal" || hyperlink.sheetId !== sheetId)
          continue;
        const address = shiftAddress(hyperlink.address);
        address !== hyperlink.address &&
          ((cells ??= { ...sheet.cells }),
          (cells[cellAddress] = { ...cell }),
          (hasChanges = true),
          address
            ? (cells[cellAddress].hyperlink = {
                ...hyperlink,
                address: address,
              })
            : delete cells[cellAddress].hyperlink);
      }
      return cells ? { ...sheet, cells: cells } : sheet;
    });
  return hasChanges
    ? {
        ...workbook,
        sheets: sheets,
        ...(reviewThreads ? { reviewThreads: reviewThreads } : {}),
      }
    : workbook;
}
function shiftFormulaWatchesOnAxisEdit(workbook, sheetId, change) {
  return workbook.formulaWatches?.length
    ? {
        ...workbook,
        formulaWatches: workbook.formulaWatches.map((formulaWatch) => {
          if (formulaWatch.orphaned || formulaWatch.anchor.sheetId !== sheetId)
            return formulaWatch;
          const parsedAddress = _t(formulaWatch.anchor.address),
            axisIndex =
              change.axis === "row" ? parsedAddress.row : parsedAddress.col;
          if (
            change.moveBefore === undefined &&
            change.remove &&
            axisIndex >= change.index &&
            axisIndex < change.index + change.count
          )
            return { ...formulaWatch, orphaned: "deleted-cells" };
          const shiftedIndex =
            change.moveBefore !== undefined
              ? mapIndexAfterMove(axisIndex, change)
              : axisIndex >= change.index + (change.remove ? change.count : 0)
                ? axisIndex + (change.remove ? -change.count : change.count)
                : axisIndex;
          if (
            shiftedIndex < 0 ||
            shiftedIndex >=
              (change.axis === "row"
                ? workbookLimits.maxRows
                : workbookLimits.maxCols)
          )
            return { ...formulaWatch, orphaned: "out-of-bounds" };
          const address =
            change.axis === "row"
              ? `${columnIndexToLetters(parsedAddress.col)}${shiftedIndex + 1}`
              : `${columnIndexToLetters(shiftedIndex)}${parsedAddress.row + 1}`;
          return address === formulaWatch.anchor.address
            ? formulaWatch
            : {
                ...formulaWatch,
                anchor: { ...formulaWatch.anchor, address: address },
              };
        }),
      }
    : workbook;
}
function createAnchoredDrawingObject(kind, range) {
  const anchorCell = _t(range.split(":")[0]);
  return {
    id: crypto.randomUUID(),
    kind: kind,
    anchor: `${columnIndexToLetters(anchorCell.col)}${anchorCell.row + 1}`,
    offsetX: 12,
    offsetY: 12,
    width: kind === "chart" ? 480 : 240,
    height: kind === "chart" ? 300 : 140,
    label: kind === "picture" ? "Picture" : kind === "chart" ? "Chart" : "Text",
    fill: "#E2EFDA",
    stroke: "#548235",
    ...(kind === "shape" ? { shape: "rectangle" } : {}),
  };
}
function buildDrawingExportNotes(exportWorkbook, exportWorksheet, sheet) {
  const notes = [];
  for (const drawingObject of sheet.objects ?? []) {
    if (drawingObject.kind !== "picture") {
      notes.push(
        `${drawingObject.label || drawingObject.kind}: native chart/shape export is not supported; retain the JSON workbook.`,
      );
      continue;
    }
    const { extension } = inspectPictureDataUrl(drawingObject.dataURL);
    drawingObject.evidence &&
      notes.push(
        `${drawingObject.label || "Picture"}: XLSX retains the image bytes and placement, but original filename, source reference and evidence provenance require native JSON; a later XLSX import records only observed media.`,
      );
    const anchorCell = _t(drawingObject.anchor),
      imageId = exportWorkbook.addImage({
        base64: drawingObject.dataURL,
        extension,
      });
    exportWorksheet.addImage(imageId, {
      tl: {
        col:
          anchorCell.col +
          drawingObject.offsetX / (sheet.columnWidths?.[anchorCell.col] ?? 84),
        row:
          anchorCell.row +
          drawingObject.offsetY / (sheet.rowHeights?.[anchorCell.row] ?? 21),
      },
      ext: { width: drawingObject.width, height: drawingObject.height },
      editAs: "oneCell",
    });
  }
  return notes;
}
function normalizeFormulaWatchAddress(address) {
  if (
    typeof address != "string" ||
    !/^\$?[A-Za-z]{1,3}\$?[1-9]\d{0,6}$/.test(address)
  )
    throw Error("A formula watch must identify one valid Excel cell address.");
  const parsedAddress = _t(address.replace(/\$/g, "").toUpperCase());
  if (
    parsedAddress.row >= workbookLimits.maxRows ||
    parsedAddress.col >= workbookLimits.maxCols
  )
    throw Error("A formula watch address is outside Excel's cell bounds.");
  return `${columnIndexToLetters(parsedAddress.col)}${parsedAddress.row + 1}`;
}
function validateFormulaWatches(formulaWatches, workbook) {
  if (formulaWatches === undefined) return [];
  if (!Array.isArray(formulaWatches) || formulaWatches.length > B2)
    throw Error("A workbook may contain at most 100 formula watches.");
  const watchIds = new Set(),
    activeCellKeys = new Set();
  return formulaWatches.map((formulaWatch) => {
    if (
      !isPlainObjectValue(formulaWatch) ||
      !HO(formulaWatch, ["id", "anchor", "orphaned"]) ||
      !isBoundedPlainText(formulaWatch.id, 100) ||
      !isPlainObjectValue(formulaWatch.anchor) ||
      !HO(formulaWatch.anchor, ["sheetId", "sheetName", "address"]) ||
      !isBoundedPlainText(formulaWatch.anchor.sheetId, 100) ||
      !isBoundedPlainText(formulaWatch.anchor.sheetName, 31)
    )
      throw Error("Invalid formula watch metadata.");
    if (watchIds.has(formulaWatch.id))
      throw Error("Formula watch IDs must be unique.");
    if (
      (watchIds.add(formulaWatch.id),
      formulaWatch.orphaned !== undefined &&
        !referenceImpactKinds.includes(formulaWatch.orphaned))
    )
      throw Error("Invalid formula watch orphan reason.");
    const anchor = formulaWatch.anchor,
      formulaWatchAddress = normalizeFormulaWatchAddress(anchor.address),
      sheet = workbook.sheets.find((sheet) => sheet.id === anchor.sheetId),
      orphaned = formulaWatch.orphaned ?? (sheet ? undefined : "missing-sheet"),
      cellKey = JSON.stringify([
        formulaWatch.anchor.sheetId,
        formulaWatchAddress,
      ]);
    if (!orphaned && activeCellKeys.has(cellKey))
      throw Error("Active formula watches must not repeat the same cell.");
    return (
      orphaned || activeCellKeys.add(cellKey),
      {
        id: formulaWatch.id,
        anchor: {
          sheetId: formulaWatch.anchor.sheetId,
          sheetName:
            !orphaned && sheet ? sheet.name : formulaWatch.anchor.sheetName,
          address: formulaWatchAddress,
        },
        ...(orphaned ? { orphaned: orphaned } : {}),
      }
    );
  });
}
function normalizeWhatIfCell(cellReference) {
  const cellFields = validateWhatIfFields(cellReference, [
      "sheetId",
      "address",
    ]),
    sheetId = requireWhatIfName(cellFields.sheetId);
  if (typeof cellFields.address != "string")
    throw Error("A what-if cell needs an Excel address.");
  const parsedAddress = _t(cellFields.address);
  return {
    sheetId: sheetId,
    address: `${columnIndexToLetters(parsedAddress.col)}${parsedAddress.row + 1}`,
  };
}
function normalizeWhatIfCellReferences(rawReferences, maxCount) {
  const cellReferences = requireWhatIfList(rawReferences, maxCount).map(
    normalizeWhatIfCell,
  );
  if (
    new Set(cellReferences.map(cellAddressKey)).size !== cellReferences.length
  )
    throw Error("What-if cell references must not repeat.");
  return cellReferences;
}
function validateWhatIfDefinition(request) {
  const fields = validateWhatIfFields(request, [
      "name",
      "kind",
      "inputs",
      "outputs",
      "scenarios",
      "orientation",
      "values",
      "rowValues",
      "columnValues",
    ]),
    definition = {
      name: requireWhatIfName(fields.name),
      inputs: normalizeWhatIfCellReferences(fields.inputs, whatIfLimits.inputs),
      outputs: normalizeWhatIfCellReferences(
        fields.outputs,
        whatIfLimits.outputs,
      ),
    };
  if (
    definition.inputs.some((input) =>
      definition.outputs.some(
        (output) => cellAddressKey(input) === cellAddressKey(output),
      ),
    )
  )
    throw Error("Changing inputs and formula results must be different cells.");
  const baseFieldNames = ["name", "kind", "inputs", "outputs"];
  if (fields.kind === "scenarios") {
    validateWhatIfFields(fields, [...baseFieldNames, "scenarios"]);
    const scenarios = requireWhatIfList(
      fields.scenarios,
      whatIfLimits.scenarios,
    ).map((scenarioRequest) => {
      const scenarioFields = validateWhatIfFields(scenarioRequest, [
          "name",
          "values",
        ]),
        values = validateWhatIfValues(
          scenarioFields.values,
          whatIfLimits.inputs,
        );
      if (values.length !== definition.inputs.length)
        throw Error(
          "Each scenario needs exactly one value per changing input.",
        );
      return { name: requireWhatIfName(scenarioFields.name), values: values };
    });
    if (
      new Set(scenarios.map((scenario) => scenario.name.toLowerCase())).size !==
      scenarios.length
    )
      throw Error("Scenario names must be unique, ignoring case.");
    return { ...definition, kind: fields.kind, scenarios: scenarios };
  }
  if (fields.kind === "one-variable") {
    if (
      (validateWhatIfFields(fields, [
        ...baseFieldNames,
        "orientation",
        "values",
      ]),
      definition.inputs.length !== 1 ||
        !["row", "column"].includes(fields.orientation))
    )
      throw Error(
        "A one-variable table needs one input and a row or column orientation.",
      );
    return {
      ...definition,
      kind: fields.kind,
      orientation: fields.orientation,
      values: validateWhatIfValues(fields.values, whatIfLimits.axisValues),
    };
  }
  if (fields.kind === "two-variable") {
    if (
      (validateWhatIfFields(fields, [
        ...baseFieldNames,
        "rowValues",
        "columnValues",
      ]),
      definition.inputs.length !== 2 || definition.outputs.length !== 1)
    )
      throw Error(
        "A two-variable table needs two different inputs and exactly one formula result.",
      );
    const rowValues = validateWhatIfValues(
        fields.rowValues,
        whatIfLimits.axisValues,
      ),
      columnValues = validateWhatIfValues(
        fields.columnValues,
        whatIfLimits.axisValues,
      );
    if (rowValues.length * columnValues.length > whatIfLimits.trials)
      throw Error("A sensitivity table supports at most 400 combinations.");
    return {
      ...definition,
      kind: fields.kind,
      rowValues: rowValues,
      columnValues: columnValues,
    };
  }
  throw Error(
    "Choose scenarios, a one-variable table or a two-variable table.",
  );
}
function validateSavedWhatIfAnalyses(whatIfAnalyses) {
  const analysisIds = new Set();
  return requireWhatIfList(whatIfAnalyses, whatIfLimits.savedAnalyses, 0).map(
    (analysis) => {
      const analysisFields = validateWhatIfFields(analysis, [
          "id",
          "reportSheetId",
          "request",
          "baseline",
          "trials",
          "assumptions",
          "sourceRevision",
          "dependencyCells",
          "workUnits",
        ]),
        analysisId = requireWhatIfName(analysisFields.id),
        reportSheetId = requireWhatIfName(analysisFields.reportSheetId),
        request = validateWhatIfDefinition(analysisFields.request);
      if (analysisIds.has(analysisId))
        throw Error("Saved what-if IDs must be unique.");
      analysisIds.add(analysisId);
      const normalizeResult = (resultRecord) => {
          const resultFields = validateWhatIfFields(resultRecord, [
              "name",
              "inputs",
              "outputs",
              "error",
            ]),
            resultName = requireWhatIfName(resultFields.name, 200),
            inputs = validateWhatIfValues(
              resultFields.inputs,
              whatIfLimits.inputs,
            ),
            outputs =
              resultFields.outputs === null
                ? null
                : validateWhatIfValues(
                    resultFields.outputs,
                    whatIfLimits.outputs,
                  ),
            resultError =
              resultFields.error === null
                ? null
                : requireWhatIfName(resultFields.error, 2e3);
          if (
            inputs.length !== request.inputs.length ||
            (outputs && outputs.length !== request.outputs.length) ||
            !!outputs == !!resultError
          )
            throw Error("Invalid saved what-if result dimensions or outcome.");
          return {
            name: resultName,
            inputs: inputs,
            outputs: outputs,
            error: resultError,
          };
        },
        baseline = normalizeResult(analysisFields.baseline),
        trials = requireWhatIfList(
          analysisFields.trials,
          whatIfLimits.trials,
        ).map(normalizeResult),
        expectedTrials = expandWhatIfScenarios(request);
      if (
        trials.length !== expectedTrials.length ||
        trials.some(
          (trial, trialIndex) =>
            trial.name !== expectedTrials[trialIndex].name ||
            JSON.stringify(trial.inputs) !==
              JSON.stringify(expectedTrials[trialIndex].inputs),
        )
      )
        throw Error(
          "Saved what-if results must match their named assumptions.",
        );
      if (baseline.name !== "Baseline")
        throw Error("Saved what-if baseline is missing.");
      const assumptionCellKeys = new Set(),
        assumptions = requireWhatIfList(
          analysisFields.assumptions,
          whatIfLimits.dependencyCells,
        ).map((assumptionRecord) => {
          const cell = validateWhatIfFields(assumptionRecord, [
              "sheetId",
              "sheetName",
              "address",
              "raw",
              "literal",
            ]),
            whatIfCell = normalizeWhatIfCell({
              sheetId: cell.sheetId,
              address: cell.address,
            });
          if (assumptionCellKeys.has(cellAddressKey(whatIfCell)))
            throw Error("Saved assumptions must not repeat a source cell.");
          if (
            (assumptionCellKeys.add(cellAddressKey(whatIfCell)),
            !(
              cell.raw === null ||
              typeof cell.raw == "boolean" ||
              isFiniteNumber(cell.raw) ||
              (typeof cell.raw == "string" && cell.raw.length <= 32767)
            ) ||
              (cell.literal !== undefined && typeof cell.literal != "boolean"))
          )
            throw Error("Invalid saved what-if assumption.");
          return {
            ...whatIfCell,
            sheetName: requireWhatIfName(cell.sheetName, 31),
            raw: cell.raw,
            ...(cell.literal === undefined ? {} : { literal: cell.literal }),
          };
        });
      if (
        [...request.inputs, ...request.outputs].some(
          (referenceCell) =>
            !assumptionCellKeys.has(cellAddressKey(referenceCell)),
        ) ||
        request.inputs.some(
          (input, inputIndex) =>
            assumptions.find(
              (assumption) =>
                cellAddressKey(assumption) === cellAddressKey(input),
            )?.raw !== baseline.inputs[inputIndex],
        )
      )
        throw Error(
          "Saved baseline inputs and dependency assumptions do not reconcile.",
        );
      for (const [fieldName, fieldLimit] of [
        ["sourceRevision", Number.MAX_SAFE_INTEGER - 1],
        ["dependencyCells", whatIfLimits.dependencyCells],
        ["workUnits", whatIfLimits.workUnits],
      ])
        if (
          !Number.isSafeInteger(analysisFields[fieldName]) ||
          analysisFields[fieldName] < 0 ||
          analysisFields[fieldName] > fieldLimit
        )
          throw Error("Invalid what-if evidence counts.");
      if (
        analysisFields.dependencyCells !== assumptions.length ||
        analysisFields.workUnits !== assumptions.length * (trials.length + 1)
      )
        throw Error("What-if work counts do not reconcile.");
      return {
        id: analysisId,
        reportSheetId: reportSheetId,
        request: request,
        baseline: baseline,
        trials: trials,
        assumptions: assumptions,
        sourceRevision: analysisFields.sourceRevision,
        dependencyCells: assumptions.length,
        workUnits: analysisFields.workUnits,
      };
    },
  );
}
function applyChartDrawingLayout(sheet, object) {
  if (object.kind === "chart" && object.chart === "sparkline") {
    const anchorCell = _t(object.anchor);
    return {
      offsetX: 2,
      offsetY: 2,
      width: columnWidthAt(sheet, anchorCell.col) - 4,
      height: rowHeightAt(sheet, anchorCell.row) - 4,
    };
  }
  return {
    offsetX: object.offsetX,
    offsetY: object.offsetY,
    width: object.width,
    height: object.height,
  };
}
function parsePrintTitles(titleText, axis, sheet) {
  if (titleText === undefined || titleText === "") return;
  if (typeof titleText != "string")
    throw Error("Print titles must be a single row or column interval.");
  const intervalMatch = (
    axis === "row"
      ? /^\$?([1-9]\d*):\$?([1-9]\d*)$/
      : /^\$?([A-Za-z]{1,3}):\$?([A-Za-z]{1,3})$/
  ).exec(titleText.trim());
  if (!intervalMatch)
    throw Error(
      axis === "row"
        ? "Repeat rows must use a range such as 1:2."
        : "Repeat columns must use a range such as A:B.",
    );
  const start =
      axis === "row"
        ? Number(intervalMatch[1]) - 1
        : _t(intervalMatch[1] + "1").col,
    end =
      axis === "row"
        ? Number(intervalMatch[2]) - 1
        : _t(intervalMatch[2] + "1").col,
    axisCount = axis === "row" ? sheet.rowCount : sheet.colCount,
    titleLimit =
      axis === "row"
        ? printRenderLimits.titleRows
        : printRenderLimits.titleColumns;
  if (
    !Number.isSafeInteger(start) ||
    !Number.isSafeInteger(end) ||
    start < 0 ||
    end < start ||
    end >= axisCount ||
    end - start + 1 > titleLimit
  )
    throw Error(
      `Print titles support at most ${titleLimit} ${axis === "row" ? "rows" : "columns"} within the worksheet.`,
    );
  return { start: start, end: end };
}
function shiftPrintSettingsOnAxisEdit(
  pageLayout,
  sheet,
  axis,
  index,
  count,
  remove,
) {
  const repeatKey = axis === "row" ? "repeatRows" : "repeatColumns",
    breakKey = axis === "row" ? "rowBreaks" : "columnBreaks",
    nextPageLayout = { ...pageLayout },
    printTitles = parsePrintTitles(pageLayout[repeatKey], axis, sheet);
  if (printTitles) {
    let { start, end } = printTitles;
    if (!remove)
      (index <= start && (start += count), index <= end && (end += count));
    else {
      const removedEnd = index + count - 1,
        remainingIntervals = [
          start < index ? { start, end: Math.min(end, index - 1) } : undefined,
          end > removedEnd
            ? {
                start: Math.max(start, removedEnd + 1) - count,
                end: end - count,
              }
            : undefined,
        ].filter(Boolean);
      if (!remainingIntervals.length)
        throw Error(
          "This edit removes every repeated print title. Clear or change print titles first.",
        );
      ((start = remainingIntervals[0].start),
        (end = remainingIntervals.at(-1).end));
    }
    ((nextPageLayout[repeatKey] = formatAxisIntervalRef({ start, end }, axis)),
      parsePrintTitles(nextPageLayout[repeatKey], axis, sheet));
  }
  const axisCount = axis === "row" ? sheet.rowCount : sheet.colCount,
    shiftedBreaks = (pageLayout[breakKey] ?? [])
      .map((breakPosition) =>
        remove
          ? breakPosition < index
            ? breakPosition
            : breakPosition >= index + count
              ? breakPosition - count
              : index
          : breakPosition >= index
            ? breakPosition + count
            : breakPosition,
      )
      .filter((shiftedPosition) => shiftedPosition > 0);
  if (shiftedBreaks.some((movedPosition) => movedPosition >= axisCount))
    throw Error(
      "This edit moves a manual print break beyond the worksheet. Remove the break first.",
    );
  return (
    shiftedBreaks.length
      ? (nextPageLayout[breakKey] = [...new Set(shiftedBreaks)].sort(
          (leftPosition, rightPosition) => leftPosition - rightPosition,
        ))
      : delete nextPageLayout[breakKey],
    nextPageLayout
  );
}
function isCanonicalCellAddress(address) {
  if (typeof address != "string") return false;
  try {
    const parsedAddress = _t(address);
    return (
      `${columnIndexToLetters(parsedAddress.col)}${parsedAddress.row + 1}` ===
      address
    );
  } catch {
    return false;
  }
}
const evidenceAddressSchema = string().refine((address) => {
    try {
      const parsedAddress = _t(address);
      return (
        `${columnIndexToLetters(parsedAddress.col)}${parsedAddress.row + 1}` ===
        address
      );
    } catch {
      return false;
    }
  }, "Invalid evidence address."),
  evidenceSourceOriginSchema = strictObject({
    id: string().min(1).max(100),
    sourceId: string().min(1).max(100),
    sourceName: string().min(1).max(255),
    sourceSheet: string().min(1).max(31),
    originalRow: number().int().min(1).max(1048576),
    kind: _enum(["file", "snapshot"]),
  }),
  evidenceCellRecordSchema = strictObject({
    address: evidenceAddressSchema,
    original: unknown().optional(),
    origin: evidenceSourceOriginSchema.optional(),
  }),
  evidenceFieldValueSchema = strictObject({
    address: evidenceAddressSchema,
    row: number().int().min(1).max(1048576),
    field: string().min(1).max(150),
    value: transformCellValueSchema,
    type: _enum(["text", "number", "boolean", "date", "blank"]),
    sources: array(evidenceAddressSchema).min(1).max(500),
    exception: boolean(),
  }),
  evidenceTransformStepSchema = strictObject({
    step: number().int().min(1).max(12),
    row: number().int().min(1).max(1048576),
    field: string().min(1).max(150),
    raw: transformCellValueSchema,
    sources: array(evidenceAddressSchema).min(1).max(500),
    reason: string().min(1).max(1e3),
  }),
  analysisOutputEvidenceSchema = strictObject({
    id: string().min(1).max(100),
    workbookId: string().min(1).max(100),
    sourceRevision: number().int().nonnegative(),
    sourceSheetId: string().min(1).max(100),
    sourceSheetName: string().min(1).max(31),
    outputSheetId: string().min(1).max(100),
    outputSheetName: string().min(1).max(31),
    createdAt: isoDateTimeSchemaFactory(),
    semantics: literal("reviewed-fields-v1"),
    request: unknown(),
    sources: array(evidenceCellRecordSchema).min(1).max(500),
    outputs: array(evidenceFieldValueSchema).min(1).max(5e3),
    exceptions: array(evidenceTransformStepSchema).max(5e3),
  });
function parseRangeAddress(rangeText) {
  const parts = rangeText.trim().split(":");
  if (parts.length > 2) return throwWithMessage(`Invalid range: ${rangeText}`);
  const firstCell = _t(parts[0]),
    lastCell = _t(parts[1] ?? parts[0]);
  return {
    firstRow: Math.min(firstCell.row, lastCell.row),
    lastRow: Math.max(firstCell.row, lastCell.row),
    firstCol: Math.min(firstCell.col, lastCell.col),
    lastCol: Math.max(firstCell.col, lastCell.col),
  };
}
function parseSelectionAreas(selectionInput) {
  const selectionParts =
    typeof selectionInput == "string"
      ? selectionInput.split(",")
      : [...selectionInput];
  if (!selectionParts.length || selectionParts.length > maxSelectionAreas)
    throw Error(`Select between 1 and ${maxSelectionAreas} rectangular areas.`);
  const normalizedAreas = selectionParts.map((areaText) => {
    const rangeAddress = parseRangeAddress(
        areaText.trim().toUpperCase().replace(/\$/g, ""),
      ),
      startAddress = `${columnIndexToLetters(rangeAddress.firstCol)}${rangeAddress.firstRow + 1}`,
      endAddress = `${columnIndexToLetters(rangeAddress.lastCol)}${rangeAddress.lastRow + 1}`;
    return startAddress === endAddress
      ? startAddress
      : `${startAddress}:${endAddress}`;
  });
  return [...new Set(normalizedAreas)];
}
function isCellAddressInRanges(selectionAreas, active) {
  const cellPosition = _t(active);
  return selectionAreas.some((selectionArea) => {
    const rangeAddress = parseRangeAddress(selectionArea);
    return (
      cellPosition.row >= rangeAddress.firstRow &&
      cellPosition.row <= rangeAddress.lastRow &&
      cellPosition.col >= rangeAddress.firstCol &&
      cellPosition.col <= rangeAddress.lastCol
    );
  });
}
function countSelectionCells(selectionAreas) {
  const rangeAddresses =
      parseSelectionAreas(selectionAreas).map(parseRangeAddress),
    rowBoundaries = [
      ...new Set(
        rangeAddresses.flatMap((rangeAddress) => [
          rangeAddress.firstRow,
          rangeAddress.lastRow + 1,
        ]),
      ),
    ].sort((rowBoundary, otherRowBoundary) => rowBoundary - otherRowBoundary);
  let cellCount = 0;
  for (
    let rowBoundaryIndex = 0;
    rowBoundaryIndex < rowBoundaries.length - 1;
    rowBoundaryIndex++
  ) {
    const columnSpans = rangeAddresses
      .filter(
        (spanningRange) =>
          spanningRange.firstRow <= rowBoundaries[rowBoundaryIndex] &&
          spanningRange.lastRow >= rowBoundaries[rowBoundaryIndex],
      )
      .map((columnSpan) => [columnSpan.firstCol, columnSpan.lastCol + 1])
      .sort((spanA, spanB) => spanA[0] - spanB[0]);
    let coveredColumnCount = 0,
      coveredEnd = -1;
    for (const [spanStart, spanEnd] of columnSpans)
      ((coveredColumnCount += Math.max(
        0,
        spanEnd - Math.max(spanStart, coveredEnd),
      )),
        (coveredEnd = Math.max(coveredEnd, spanEnd)));
    cellCount +=
      coveredColumnCount *
      (rowBoundaries[rowBoundaryIndex + 1] - rowBoundaries[rowBoundaryIndex]);
  }
  return cellCount;
}
function enumerateSelectionCells(
  selection,
  maxBlankScan = workbookLimits.maxPatches,
) {
  if (
    !Number.isSafeInteger(maxBlankScan) ||
    maxBlankScan < 1 ||
    maxBlankScan > workbookLimits.maxCells
  )
    throw Error("Invalid selection enumeration limit.");
  const selectionAreas = parseSelectionAreas(selection),
    cellCount = countSelectionCells(selectionAreas);
  if (cellCount > maxBlankScan)
    throw Error(
      `Select at most ${maxBlankScan.toLocaleString()} cells for this operation. The current selection contains ${cellCount.toLocaleString()} cells.`,
    );
  const cellKeys = new Set();
  for (const selectionArea of selectionAreas) {
    const rangeAddress = parseRangeAddress(selectionArea);
    for (
      let firstRow = rangeAddress.firstRow;
      firstRow <= rangeAddress.lastRow;
      firstRow++
    )
      for (
        let firstCol = rangeAddress.firstCol;
        firstCol <= rangeAddress.lastCol;
        firstCol++
      )
        cellKeys.add(`${columnIndexToLetters(firstCol)}${firstRow + 1}`);
  }
  return [...cellKeys];
}
function pQ(workbook, sourceSheetId, sortRangeAddress, sheetId, targetRange) {
  const rangeAddress = parseRangeAddress(sortRangeAddress),
    targetAddress = parseRangeAddress(targetRange),
    addressInRange = (addressRange, cellAddress) => {
      const position = _t(cellAddress);
      return (
        position.row >= addressRange.firstRow &&
        position.row <= addressRange.lastRow &&
        position.col >= addressRange.firstCol &&
        position.col <= addressRange.lastCol
      );
    },
    changes = [],
    sheet = workbook.sheets.find((sheet) => sheet.id === sheetId);
  return {
    watches: workbook.formulaWatches?.map((formulaWatch) => {
      if (formulaWatch.orphaned) return formulaWatch;
      const before = `${formulaWatch.anchor.sheetName}!${formulaWatch.anchor.address}`;
      if (
        formulaWatch.anchor.sheetId === sourceSheetId &&
        addressInRange(rangeAddress, formulaWatch.anchor.address)
      ) {
        const anchorPosition = _t(formulaWatch.anchor.address),
          address = `${columnIndexToLetters(anchorPosition.col + targetAddress.firstCol - rangeAddress.firstCol)}${anchorPosition.row + targetAddress.firstRow - rangeAddress.firstRow + 1}`,
          after = `${sheet.name}!${address}`;
        return before === after && sourceSheetId === sheetId
          ? formulaWatch
          : (changes.push({
              kind: "formula-watch",
              label: formulaWatch.id,
              before: before,
              after: after,
            }),
            {
              ...formulaWatch,
              anchor: {
                sheetId: sheetId,
                sheetName: sheet.name,
                address: address,
              },
            });
      }
      return formulaWatch.anchor.sheetId === sheetId &&
        addressInRange(targetAddress, formulaWatch.anchor.address)
        ? (changes.push({
            kind: "formula-watch",
            label: formulaWatch.id,
            before: before,
            after: `${before} — watched source overwritten`,
          }),
          { ...formulaWatch, orphaned: "overwritten" })
        : formulaWatch;
    }),
    changes: changes,
  };
}
function validateFreezePanes(
  nextSheet,
  sheet,
  gridLimits = { rows: sheet.rowCount, cols: sheet.colCount },
) {
  validateSheetDefaultSizes(nextSheet, sheet);
  const normalizedSplitPanes = validateSplitPaneState(
    nextSheet.splitPanes,
    gridLimits,
  );
  if (normalizedSplitPanes && (nextSheet.frozenRows || nextSheet.frozenCols))
    throw Error("Split panes and frozen panes cannot be enabled together.");
  normalizedSplitPanes && (sheet.splitPanes = normalizedSplitPanes);
  const hiddenIntervals = normalizeHiddenIntervals(
      nextSheet.hiddenRows,
      gridLimits.rows,
    ),
    hiddenColIntervals = normalizeHiddenIntervals(
      nextSheet.hiddenCols,
      gridLimits.cols,
    ),
    sheetWithHidden = {
      ...sheet,
      hiddenRows: hiddenIntervals,
      hiddenCols: hiddenColIntervals,
    },
    axisVisibilityIndex = createAxisVisibilityIndex(
      sheet.rowCount,
      cn(sheetWithHidden, "row"),
    ),
    colVisibilityIndex = createAxisVisibilityIndex(
      sheet.colCount,
      cn(sheetWithHidden, "column"),
    );
  if (!axisVisibilityIndex.count || !colVisibilityIndex.count)
    throw Error("At least one row and one column must remain visible.");
  const frozenRowCount =
      nextSheet.frozenRows === undefined ? 0 : nextSheet.frozenRows,
    frozenColCount =
      nextSheet.frozenCols === undefined ? 0 : nextSheet.frozenCols;
  if (
    !Number.isInteger(frozenRowCount) ||
    !Number.isInteger(frozenColCount) ||
    frozenRowCount < 0 ||
    frozenColCount < 0 ||
    frozenRowCount > gridLimits.rows ||
    frozenColCount > gridLimits.cols
  )
    throw Error(
      "Frozen boundaries must be valid leading row and column counts.",
    );
  const frozenVisibleRows = axisVisibilityIndex.countBefore(frozenRowCount),
    frozenVisibleCols = colVisibilityIndex.countBefore(frozenColCount);
  if (
    frozenVisibleRows > worksheetCapacityLimits.maxFrozenVisibleRows ||
    frozenVisibleCols > worksheetCapacityLimits.maxFrozenVisibleCols
  )
    throw Error(
      "Freeze panes supports at most 100 visible rows and 50 visible columns.",
    );
  if (
    frozenVisibleRows === axisVisibilityIndex.count ||
    frozenVisibleCols === colVisibilityIndex.count
  )
    throw Error("Freeze panes must leave an unfrozen visible row and column.");
  for (const mergeAddress of sheet.merges ?? []) {
    const rangeAddress = parseRangeAddress(mergeAddress);
    if (
      (rangeAddress.firstRow < frozenRowCount &&
        rangeAddress.lastRow >= frozenRowCount) ||
      (rangeAddress.firstCol < frozenColCount &&
        rangeAddress.lastCol >= frozenColCount)
    )
      throw Error(
        "Freeze panes cannot cross a merged range. Unmerge those cells first.",
      );
  }
  (nextSheet.hiddenRows !== undefined && (sheet.hiddenRows = hiddenIntervals),
    nextSheet.hiddenCols !== undefined &&
      (sheet.hiddenCols = hiddenColIntervals),
    nextSheet.frozenRows !== undefined && (sheet.frozenRows = frozenRowCount),
    nextSheet.frozenCols !== undefined && (sheet.frozenCols = frozenColCount));
}
function updateWorksheetById(workbook, sheetId, updateSheet) {
  const sheet = workbook.sheets.find((sheet) => sheet.id === sheetId);
  if (!sheet) throw Error("Worksheet not found.");
  assertSheetUnprotected(sheet);
  const updatedSheet = updateSheet(structuredClone(sheet));
  return (
    validateFreezePanes(updatedSheet, updatedSheet),
    {
      ...workbook,
      sheets: workbook.sheets.map((sheet) =>
        sheet.id === sheetId ? updatedSheet : sheet,
      ),
    }
  );
}
function setHiddenAxisRange(workbook, sheetId, axis, start, end, isHidden) {
  if ((assertWorksheetAxis(axis), typeof isHidden != "boolean"))
    throw Error("Hidden state must be boolean.");
  return updateWorksheetById(workbook, sheetId, (sheetDraft) => {
    const axisCount = axisCountOf(sheetDraft, axis);
    normalizeHiddenIntervals([{ start: start, end: end }], axisCount);
    const hiddenAxisIntervals = readHiddenAxisIntervals(sheetDraft, axis),
      newInterval = { start: start, end: end },
      remainingIntervals = isHidden
        ? hiddenAxisIntervals.filter((hiddenAxisInterval) =>
            hiddenAxisInterval.end < newInterval.start - 1 ||
            hiddenAxisInterval.start > newInterval.end + 1
              ? true
              : ((newInterval.start = Math.min(
                  newInterval.start,
                  hiddenAxisInterval.start,
                )),
                (newInterval.end = Math.max(
                  newInterval.end,
                  hiddenAxisInterval.end,
                )),
                false),
          )
        : hiddenAxisIntervals,
      updatedIntervals = isHidden
        ? [...remainingIntervals, newInterval]
        : hiddenAxisIntervals.flatMap((hiddenAxisInterval) =>
            hiddenAxisInterval.end < start || hiddenAxisInterval.start > end
              ? [hiddenAxisInterval]
              : [
                  ...(hiddenAxisInterval.start < start
                    ? [{ start: hiddenAxisInterval.start, end: start - 1 }]
                    : []),
                  ...(hiddenAxisInterval.end > end
                    ? [{ start: end + 1, end: hiddenAxisInterval.end }]
                    : []),
                ],
          );
    return (
      (sheetDraft[hiddenKeyForAxis(axis)] = normalizeHiddenIntervals(
        updatedIntervals,
        axisCount,
      )),
      sheetDraft
    );
  });
}
function yQ(workbook, sheetId, axis) {
  return (
    assertWorksheetAxis(axis),
    updateWorksheetById(
      workbook,
      sheetId,
      (sheetDraft) => (delete sheetDraft[hiddenKeyForAxis(axis)], sheetDraft),
    )
  );
}
function setFreezePanes(workbook, sheetId, frozenRows, frozenCols) {
  return updateWorksheetById(
    workbook,
    sheetId,
    (sheetDraft) => (
      delete sheetDraft.splitPanes,
      { ...sheetDraft, frozenRows: frozenRows, frozenCols: frozenCols }
    ),
  );
}
function setSplitPanes(workbook, sheetId, splitPanes) {
  return updateWorksheetById(
    workbook,
    sheetId,
    (sheetDraft) => (
      splitPanes === undefined
        ? delete sheetDraft.splitPanes
        : ((sheetDraft.splitPanes = validateSplitPaneState(splitPanes, {
            rows: sheetDraft.rowCount,
            cols: sheetDraft.colCount,
          })),
          (sheetDraft.frozenRows = 0),
          (sheetDraft.frozenCols = 0)),
      sheetDraft
    ),
  );
}
function setFreezePanesAtCell(workbook, sheetId, cellAddress) {
  const cellPosition = _t(cellAddress);
  return setFreezePanes(workbook, sheetId, cellPosition.row, cellPosition.col);
}
function applyAxisStructureEdit(sheet, axis, start, count, remove) {
  assertWorksheetAxis(axis);
  const axisCount = axisCountOf(sheet, axis);
  if (
    !Number.isInteger(start) ||
    start < 0 ||
    !Number.isInteger(count) ||
    count < 1 ||
    count > 1e4 ||
    start + count > axisCount ||
    typeof remove != "boolean"
  )
    throw Error("Choose a valid structural edit of 1–10,000 rows or columns.");
  const shiftedIntervals = readHiddenAxisIntervals(sheet, axis).flatMap(
      (hiddenInterval) => {
        if (!remove) {
          if (hiddenInterval.end < start) return [hiddenInterval];
          if (hiddenInterval.end + count >= axisCount)
            throw Error(
              "Insertion would discard hidden-axis metadata at the worksheet boundary.",
            );
          return [
            {
              start:
                hiddenInterval.start >= start
                  ? hiddenInterval.start + count
                  : hiddenInterval.start,
              end: hiddenInterval.end + count,
            },
          ];
        }
        const lastRemovedIndex = start + count - 1;
        return hiddenInterval.end < start
          ? [hiddenInterval]
          : hiddenInterval.start > lastRemovedIndex
            ? [
                {
                  start: hiddenInterval.start - count,
                  end: hiddenInterval.end - count,
                },
              ]
            : [
                ...(hiddenInterval.start < start
                  ? [{ start: hiddenInterval.start, end: start - 1 }]
                  : []),
                ...(hiddenInterval.end > lastRemovedIndex
                  ? [{ start: start, end: hiddenInterval.end - count }]
                  : []),
              ];
      },
    ),
    hiddenKey = hiddenKeyForAxis(axis),
    frozenKey = axis === "row" ? "frozenRows" : "frozenCols",
    currentFrozenCount = sheet[frozenKey] ?? 0,
    nextFrozenCount = remove
      ? currentFrozenCount -
        Math.max(0, Math.min(count, currentFrozenCount - start))
      : currentFrozenCount + (start < currentFrozenCount ? count : 0),
    editedSheet = {
      ...sheet,
      [hiddenKey]: normalizeHiddenIntervals(shiftedIntervals, axisCount),
      [frozenKey]: nextFrozenCount,
      ...(sheet.outlineGroups
        ? {
            outlineGroups: remapOutlineGroupsForEdit(
              sheet,
              axis,
              start,
              count,
              remove,
            ),
          }
        : {}),
    };
  return (
    validateFreezePanes(editedSheet, editedSheet),
    {
      [hiddenKey]: editedSheet[hiddenKey],
      [frozenKey]: nextFrozenCount,
      ...(sheet.outlineGroups
        ? { outlineGroups: editedSheet.outlineGroups }
        : {}),
    }
  );
}
function setWorksheetReadOnly(workbook, sheetId, mutateSheet, options = {}) {
  if (options.readOnly !== undefined && typeof options.readOnly != "boolean")
    throw Error("Read-only state must be boolean.");
  if (options.readOnly) throw Error("This workbook is read-only.");
  const sheet = workbook.sheets.find((sheet) => sheet.id === sheetId);
  if (!sheet) throw Error("Worksheet not found.");
  assertSheetUnprotected(sheet);
  const sheetClone = structuredClone(sheet);
  return (
    mutateSheet(sheetClone),
    validateOutlineGroups(sheetClone, sheetClone),
    validateFreezePanes(sheetClone, sheetClone),
    JSON.stringify(sheetClone) === JSON.stringify(sheet)
      ? workbook
      : {
          ...workbook,
          sheets: workbook.sheets.map((sheet) =>
            sheet.id === sheetId ? sheetClone : sheet,
          ),
        }
  );
}
function EQ(workbook, sheetId, axis, start, end, options) {
  return (
    assertWorksheetAxis(axis),
    setWorksheetReadOnly(
      workbook,
      sheetId,
      (sheetDraft) => {
        sheetDraft.outlineGroups = [
          ...(sheetDraft.outlineGroups ?? []),
          {
            id: crypto.randomUUID(),
            axis: axis,
            start: start,
            end: end,
            collapsed: false,
          },
        ];
      },
      options,
    )
  );
}
function removeOutlineGroup(workbook, sheetId, outlineGroupId, options) {
  return setWorksheetReadOnly(
    workbook,
    sheetId,
    (sheetDraft) => {
      if (
        !sheetDraft.outlineGroups?.some(
          (outlineGroup) => outlineGroup.id === outlineGroupId,
        )
      )
        throw Error("Outline group not found.");
      sheetDraft.outlineGroups = sheetDraft.outlineGroups.filter(
        (outlineGroup) => outlineGroup.id !== outlineGroupId,
      );
    },
    options,
  );
}
function setOutlineGroupCollapsed(
  workbook,
  sheetId,
  outlineGroupId,
  isCollapsed,
  options,
) {
  if (typeof isCollapsed != "boolean")
    throw Error("Collapsed state must be boolean.");
  return setWorksheetReadOnly(
    workbook,
    sheetId,
    (sheetDraft) => {
      const targetGroup = sheetDraft.outlineGroups?.find(
        (outlineGroup) => outlineGroup.id === outlineGroupId,
      );
      if (!targetGroup) throw Error("Outline group not found.");
      targetGroup.collapsed = isCollapsed;
    },
    options,
  );
}
function setOutlineGroupLevel(workbook, sheetId, axis, outlineLevel, options) {
  return (
    assertWorksheetAxis(axis),
    setWorksheetReadOnly(
      workbook,
      sheetId,
      (sheet) => {
        const axisGroups = addOutlineGroup(sheet, axis),
          levelLimit =
            Math.max(0, ...axisGroups.map((group) => group.depth)) + 1;
        if (
          !Number.isInteger(outlineLevel) ||
          outlineLevel < 1 ||
          outlineLevel > levelLimit
        )
          throw Error("Choose an available outline level.");
        const depthById = new Map(
          axisGroups.map((entry) => [entry.id, entry.depth]),
        );
        sheet.outlineGroups = sheet.outlineGroups?.map((outlineGroup) =>
          outlineGroup.axis === axis
            ? {
                ...outlineGroup,
                collapsed: depthById.get(outlineGroup.id) >= outlineLevel,
              }
            : outlineGroup,
        );
      },
      options,
    )
  );
}
function clearAxisOutlineGroups(workbook, sheetId, axis, options) {
  return (
    assertWorksheetAxis(axis),
    setWorksheetReadOnly(
      workbook,
      sheetId,
      (sheetDraft) => {
        sheetDraft.outlineGroups = sheetDraft.outlineGroups?.filter(
          (outlineGroup) => outlineGroup.axis !== axis,
        );
      },
      options,
    )
  );
}
function normalizeDataValidationRules(validations) {
  if (
    !Array.isArray(validations) ||
    validations.length > validationLimits.rulesPerSheet
  )
    throw Error("At most 100 validation rules are supported.");
  const seenIds = new Set();
  return validations.map((validation) => {
    if (
      !isNonArrayObject(validation) ||
      ![
        "list",
        "number",
        "integer",
        "textLength",
        "date",
        "time",
        "custom",
      ].includes(String(validation.kind)) ||
      ![
        "between",
        "notBetween",
        "greater",
        "greaterEqual",
        "less",
        "lessEqual",
        "equal",
        "notEqual",
      ].includes(String(validation.operator)) ||
      typeof validation.allowBlank != "boolean" ||
      !Array.isArray(validation.list) ||
      validation.list.length > validationLimits.listItems ||
      typeof validation.min != "number" ||
      !Number.isFinite(validation.min) ||
      typeof validation.max != "number" ||
      !Number.isFinite(validation.max)
    )
      throw Error("Invalid data validation rule.");
    const ruleId = assertBoundedValidationText(validation.id, 100, "ID");
    if (!ruleId.trim() || seenIds.has(ruleId))
      throw Error("Validation rule IDs must be nonempty and unique.");
    seenIds.add(ruleId);
    const rangeAddress = parseRangeAddress(
        assertBoundedValidationText(validation.range, 40, "range"),
      ),
      rule = {
        id: ruleId,
        range: `${columnIndexToLetters(rangeAddress.firstCol)}${rangeAddress.firstRow + 1}:${columnIndexToLetters(rangeAddress.lastCol)}${rangeAddress.lastRow + 1}`,
        kind: validation.kind,
        operator: validation.operator,
        min: validation.min,
        max: validation.max,
        list: validation.list.map((listItem) =>
          assertBoundedValidationText(listItem, 1e3, "list item"),
        ),
        allowBlank: validation.allowBlank,
        message: assertBoundedValidationText(
          validation.message,
          500,
          "error message",
        ),
      };
    if (
      rule.kind === "list" &&
      (!rule.list.length || rule.list.some((entry) => !entry.length))
    )
      throw Error("A list rule needs nonempty values.");
    if (
      !["list", "custom"].includes(rule.kind) &&
      ["between", "notBetween"].includes(rule.operator) &&
      rule.min > rule.max
    )
      throw Error("Minimum cannot exceed maximum.");
    if (
      rule.kind === "time" &&
      (rule.min < 0 || rule.min >= 1 || rule.max < 0 || rule.max >= 1)
    )
      throw Error(
        "Time bounds must be fractions from 0 up to, but excluding, 1.",
      );
    if (
      validation.formula !== undefined &&
      ((rule.formula = assertBoundedValidationText(
        validation.formula,
        8192,
        "formula",
      )),
      !rule.formula.startsWith("=") || rule.formula.length < 2)
    )
      throw Error("Start the custom validation formula with =.");
    if (rule.kind === "custom" && !rule.formula)
      throw Error(
        "A custom rule requires a formula relative to the range's top-left cell.",
      );
    if (validation.errorStyle !== undefined) {
      if (
        !["stop", "warning", "information"].includes(
          String(validation.errorStyle),
        )
      )
        throw Error("Invalid validation error style.");
      rule.errorStyle = validation.errorStyle;
    }
    for (const settingName of ["showInputMessage", "showDropdown"])
      if (validation[settingName] !== undefined) {
        if (typeof validation[settingName] != "boolean")
          throw Error("Invalid validation display setting.");
        rule[settingName] = validation[settingName];
      }
    for (const [fieldName, maxLength] of [
      ["inputTitle", 100],
      ["inputMessage", 500],
      ["errorTitle", 100],
    ])
      validation[fieldName] !== undefined &&
        (rule[fieldName] = assertBoundedValidationText(
          validation[fieldName],
          maxLength,
          fieldName,
        ));
    return rule;
  });
}
function findDataValidationsInRange(activeWorkbook, sheetId, cellAddress) {
  const cellPosition = _t(cellAddress);
  return (
    activeWorkbook.sheets.find((sheet) => sheet.id === sheetId)?.validations ??
    []
  ).filter((validation) => {
    const rangeAddress = parseRangeAddress(validation.range);
    return (
      cellPosition.row >= rangeAddress.firstRow &&
      cellPosition.row <= rangeAddress.lastRow &&
      cellPosition.col >= rangeAddress.firstCol &&
      cellPosition.col <= rangeAddress.lastCol
    );
  });
}
function normalizeConditionalRules(rules) {
  if (!Array.isArray(rules) || rules.length > 100)
    throw Error("At most 100 conditional formatting rules are supported.");
  const seenRuleIds = new Set(),
    usedPriorities = new Set();
  return rules.map((rule) => {
    if (!isNonArrayObject(rule) || !conditionalRuleKinds.includes(rule.kind))
      throw Error("Invalid conditional rule.");
    const ruleId = assertBoundedString(rule.id, 100, "conditional rule ID");
    if (!ruleId.trim() || seenRuleIds.has(ruleId))
      throw Error("Conditional rule IDs must be nonempty and unique.");
    seenRuleIds.add(ruleId);
    const rangeAddress = parseRangeAddress(
        assertBoundedString(rule.range, 40, "conditional range"),
      ),
      normalizedRule = {
        id: ruleId,
        range: `${columnIndexToLetters(rangeAddress.firstCol)}${rangeAddress.firstRow + 1}:${columnIndexToLetters(rangeAddress.lastCol)}${rangeAddress.lastRow + 1}`,
        kind: rule.kind,
        value: assertBoundedString(rule.value, 1e3, "conditional value"),
        second: assertBoundedString(
          rule.second,
          1e3,
          "conditional second value",
        ),
        fill: assertSixDigitHexColor(rule.fill),
        color: assertSixDigitHexColor(rule.color),
      };
    if (rule.priority !== undefined) {
      if (
        !Number.isInteger(rule.priority) ||
        Number(rule.priority) < 1 ||
        Number(rule.priority) > 100 ||
        usedPriorities.has(Number(rule.priority))
      )
        throw Error(
          "Conditional priorities must be unique integers from 1 to 100.",
        );
      ((normalizedRule.priority = Number(rule.priority)),
        usedPriorities.add(normalizedRule.priority));
    }
    if (rule.stopIfTrue !== undefined) {
      if (typeof rule.stopIfTrue != "boolean")
        throw Error("Stop if true must be boolean.");
      if (
        rule.stopIfTrue &&
        ["colorScale", "dataBar", "iconSet"].includes(normalizedRule.kind)
      )
        throw Error(
          "Stop if true is not supported for color scales, data bars, or icon sets.",
        );
      normalizedRule.stopIfTrue = rule.stopIfTrue;
    }
    if (
      (rule.formula !== undefined &&
        (normalizedRule.formula = assertBoundedString(
          rule.formula,
          4096,
          "conditional formula",
        )),
      normalizedRule.kind === "formula" &&
        (!normalizedRule.formula?.startsWith("=") ||
          !normalizedRule.formula.slice(1).trim()))
    )
      throw Error("A formula rule requires a formula beginning with =.");
    for (const colorKey of ["minColor", "maxColor"])
      rule[colorKey] !== undefined &&
        (normalizedRule[colorKey] = assertSixDigitHexColor(rule[colorKey]));
    if (rule.thresholdType !== undefined) {
      if (rule.thresholdType !== "percent" && rule.thresholdType !== "number")
        throw Error("Invalid icon threshold type.");
      normalizedRule.thresholdType = rule.thresholdType;
    }
    if (rule.iconThresholds !== undefined) {
      if (
        !Array.isArray(rule.iconThresholds) ||
        rule.iconThresholds.length !== 2 ||
        rule.iconThresholds.some(
          (iconThreshold) =>
            typeof iconThreshold != "number" || !Number.isFinite(iconThreshold),
        ) ||
        rule.iconThresholds[0] >= rule.iconThresholds[1] ||
        ((normalizedRule.thresholdType ?? "percent") === "percent" &&
          (rule.iconThresholds[0] < 0 || rule.iconThresholds[1] > 100))
      )
        throw Error(
          "Choose two increasing icon thresholds; percentages must be between 0 and 100.",
        );
      normalizedRule.iconThresholds = [...rule.iconThresholds];
    }
    return normalizedRule;
  });
}
function evaluateConditionalFormatRules(sheet, calculator, cells) {
  const errorsByKey = new Map(),
    cellResultCache = new Map();
  let inspectedCellCount = 0,
    formulaEvaluationCount = 0;
  const recordError = (ruleId, address, message) => {
    const errorRecord = { ruleId: ruleId, address: address, message: message };
    return (errorsByKey.set(`${ruleId}:${message}`, errorRecord), errorRecord);
  };
  let discoveryError = "";
  if (
    !cells &&
    (sheet.conditionalFormats ?? []).some((conditionalFormat) =>
      ["duplicate", "unique", "colorScale", "dataBar", "iconSet"].includes(
        conditionalFormat.kind,
      ),
    )
  )
    try {
      cells = vx(sheet, calculator);
    } catch (error) {
      discoveryError =
        error instanceof Error
          ? error.message
          : "Array values could not be discovered for conditional statistics.";
    }
  const cellEntries = Object.keys(cells ?? sheet.cells).map((address) => ({
      address: address,
      point: _t(address),
    })),
    preparedRules = sortConditionalRulesByPriority(
      sheet.conditionalFormats ?? [],
    ).map((rule) => {
      const rangeAddress = parseRangeAddress(rule.range),
        counts = new Map();
      let min = 1 / 0,
        max = -1 / 0,
        failure = "";
      if (
        (["greater", "less", "between"].includes(rule.kind) &&
          (!rule.value.trim() ||
            !Number.isFinite(Number(rule.value)) ||
            (rule.kind === "between" &&
              (!rule.second.trim() ||
                !Number.isFinite(Number(rule.second)) ||
                Number(rule.value) > Number(rule.second)))) &&
          (failure =
            "Numeric rule operands must be finite and minimum cannot exceed maximum."),
        ["duplicate", "unique", "colorScale", "dataBar", "iconSet"].includes(
          rule.kind,
        ))
      ) {
        discoveryError && (failure = discoveryError);
        for (const { address, point } of cellEntries) {
          if (++inspectedCellCount > 1e6) {
            failure =
              "Conditional statistics exceed 1,000,000 inspected cells; narrow the rule ranges or remove rules.";
            break;
          }
          if (!isCellInsideRange(rangeAddress, point)) continue;
          const cellValue = calculator(sheet.id, address);
          if (
            (typeof cellValue == "number" &&
              Number.isFinite(cellValue) &&
              ((min = Math.min(min, cellValue)),
              (max = Math.max(max, cellValue))),
            cellValue !== null && cellValue !== "")
          ) {
            const valueKey = String(cellValue).toLowerCase();
            counts.set(valueKey, (counts.get(valueKey) ?? 0) + 1);
          }
        }
      }
      return (
        failure && recordError(rule.id, "", failure),
        {
          rule: rule,
          bounds: rangeAddress,
          counts: counts,
          min: min,
          max: max,
          failure: failure,
        }
      );
    }),
    evaluateCell = (cellAddress) => {
      if (cellResultCache.has(cellAddress))
        return structuredClone(cellResultCache.get(cellAddress));
      const cellPoint = _t(cellAddress),
        cellResult = { style: {}, errors: [] },
        currentValue = calculator(sheet.id, cellAddress);
      for (const { rule, bounds, counts, min, max, failure } of preparedRules) {
        if (!isCellInsideRange(bounds, cellPoint)) continue;
        if (failure) {
          cellResult.errors.push(recordError(rule.id, cellAddress, failure));
          continue;
        }
        let isMatch = false,
          ruleStyle = {};
        if (rule.kind === "formula") {
          let formulaResult;
          try {
            if (++formulaEvaluationCount > 5e3)
              throw Error(
                "Conditional formula evaluation exceeds 5,000 cells; narrow the visible or preview range.",
              );
            if (!calculator.evaluateFormula)
              throw Error(
                "This calculator cannot evaluate conditional formulas.",
              );
            const shiftedFormula = shiftFormulaReferences(
              rule.formula,
              cellPoint.row - bounds.firstRow,
              cellPoint.col - bounds.firstCol,
            );
            if (
              ((formulaResult = calculator.evaluateFormula(
                sheet.id,
                cellAddress,
                shiftedFormula,
              )),
              isErrorValueText(formulaResult))
            )
              throw Error(`Formula returned ${formulaResult}.`);
            if (
              typeof formulaResult != "number" &&
              typeof formulaResult != "boolean"
            )
              throw Error("Formula must return a scalar boolean or number.");
            isMatch = !!formulaResult;
          } catch (error) {
            cellResult.errors.push(
              recordError(
                rule.id,
                cellAddress,
                error instanceof Error
                  ? error.message
                  : "Formula evaluation failed.",
              ),
            );
            continue;
          }
        } else if (rule.kind === "blank" || rule.kind === "notBlank")
          isMatch =
            (currentValue === null || currentValue === "") ==
            (rule.kind === "blank");
        else if (rule.kind === "error" || rule.kind === "notError")
          try {
            if (++formulaEvaluationCount > 5e3)
              throw Error(
                "Conditional formula evaluation exceeds 5,000 cells; narrow the visible or preview range.",
              );
            if (!calculator.evaluateFormula)
              throw Error(
                "This calculator cannot distinguish error values from literal text.",
              );
            const errorCheckResult = calculator.evaluateFormula(
              sheet.id,
              cellAddress,
              "=ISERROR(" + cellAddress + ")",
            );
            if (typeof errorCheckResult != "boolean")
              throw Error(
                "Error rule could not evaluate: " +
                  String(errorCheckResult) +
                  ".",
              );
            isMatch = errorCheckResult === (rule.kind === "error");
          } catch (error) {
            cellResult.errors.push(
              recordError(
                rule.id,
                cellAddress,
                error instanceof Error
                  ? error.message
                  : "Error rule evaluation failed.",
              ),
            );
            continue;
          }
        else if (rule.kind === "duplicate" || rule.kind === "unique")
          isMatch =
            currentValue !== null &&
            currentValue !== "" &&
            (counts.get(String(currentValue).toLowerCase()) ?? 0) > 1 ==
              (rule.kind === "duplicate");
        else if (rule.kind === "contains")
          isMatch = String(currentValue ?? "")
            .toLowerCase()
            .includes(rule.value.toLowerCase());
        else if (rule.kind === "equal")
          isMatch =
            String(currentValue ?? "").toLowerCase() ===
            rule.value.toLowerCase();
        else if (typeof currentValue == "number") {
          if (rule.kind === "greater")
            isMatch = currentValue > Number(rule.value);
          else if (rule.kind === "less")
            isMatch = currentValue < Number(rule.value);
          else if (rule.kind === "between")
            isMatch =
              currentValue >= Number(rule.value) &&
              currentValue <= Number(rule.second);
          else if (Number.isFinite(min) && Number.isFinite(max)) {
            const scaleRatio =
              max === min
                ? 0.5
                : Math.max(0, Math.min(1, (currentValue - min) / (max - min)));
            if (rule.kind === "colorScale") {
              const startColor = rule.minColor ?? "#f8e99b",
                endColor = rule.maxColor ?? "#63be7b";
              ((ruleStyle.fill =
                "#" +
                [1, 3, 5]
                  .map((channelOffset) =>
                    Math.round(
                      parseInt(
                        startColor.slice(channelOffset, channelOffset + 2),
                        16,
                      ) +
                        (parseInt(
                          endColor.slice(channelOffset, channelOffset + 2),
                          16,
                        ) -
                          parseInt(
                            startColor.slice(channelOffset, channelOffset + 2),
                            16,
                          )) *
                          scaleRatio,
                    )
                      .toString(16)
                      .padStart(2, "0"),
                  )
                  .join("")),
                (isMatch = true));
            } else if (rule.kind === "dataBar" && !cellResult.dataBar) {
              const domainMin = Math.min(0, min),
                domainMax = Math.max(0, max),
                span = domainMax - domainMin || 1,
                axis = (-domainMin / span) * 100,
                valuePosition = ((currentValue - domainMin) / span) * 100;
              ((cellResult.dataBar = {
                start: Math.min(axis, valuePosition),
                width: Math.abs(valuePosition - axis),
                axis: axis,
                color: currentValue < 0 ? "#d13438" : rule.fill,
                negative: currentValue < 0,
              }),
                (isMatch = true));
            } else if (rule.kind === "iconSet" && !cellResult.icon) {
              const thresholds = rule.iconThresholds ?? [33, 67],
                comparisonValue =
                  (rule.thresholdType ?? "percent") === "percent"
                    ? scaleRatio * 100
                    : currentValue;
              ((cellResult.icon =
                comparisonValue < thresholds[0]
                  ? { symbol: "▼", color: "#c42b1c", label: "Low" }
                  : comparisonValue < thresholds[1]
                    ? { symbol: "●", color: "#8a6100", label: "Middle" }
                    : { symbol: "▲", color: "#107c41", label: "High" }),
                (isMatch = true));
            }
          }
        }
        if (
          isMatch &&
          (["colorScale", "dataBar", "iconSet"].includes(rule.kind) ||
            (ruleStyle = { fill: rule.fill, color: rule.color }),
          (cellResult.style = { ...ruleStyle, ...cellResult.style }),
          rule.stopIfTrue)
        )
          break;
      }
      return (
        cellResultCache.size < 1e4 &&
          cellResultCache.set(cellAddress, structuredClone(cellResult)),
        cellResult
      );
    };
  return (
    (evaluateCell.errors = () =>
      [...errorsByKey.values()].map((recordedError) => ({ ...recordedError }))),
    evaluateCell
  );
}
const normalizeWorksheetToolText = (text, maxLength = 200) =>
    typeof text == "string" && text.length <= maxLength
      ? text
      : failWithMessage("Invalid worksheet tool text."),
  CE = (rangeText) => {
    const rangeAddress = parseRangeAddress(rangeText);
    return `${columnIndexToLetters(rangeAddress.firstCol)}${rangeAddress.firstRow + 1}:${columnIndexToLetters(rangeAddress.lastCol)}${rangeAddress.lastRow + 1}`;
  },
  doesRangeOverlap = (rangeText, otherRangeText) => {
    const rangeAddress = parseRangeAddress(rangeText),
      otherRangeAddress = parseRangeAddress(otherRangeText);
    return (
      rangeAddress.firstRow <= otherRangeAddress.lastRow &&
      rangeAddress.lastRow >= otherRangeAddress.firstRow &&
      rangeAddress.firstCol <= otherRangeAddress.lastCol &&
      rangeAddress.lastCol >= otherRangeAddress.firstCol
    );
  },
  lD = (rangeText, cellAddress) => {
    const rangeAddress = parseRangeAddress(rangeText),
      cellPoint = _t(cellAddress);
    return (
      cellPoint.row >= rangeAddress.firstRow &&
      cellPoint.row <= rangeAddress.lastRow &&
      cellPoint.col >= rangeAddress.firstCol &&
      cellPoint.col <= rangeAddress.lastCol
    );
  };
function validateWorksheetFeatures(sheet, targetSheet) {
  if (isNonArrayObject(sheet)) {
    if (sheet.merges !== undefined) {
      (!Array.isArray(sheet.merges) || sheet.merges.length > 1e3) &&
        failWithMessage("At most1,000 merged ranges are supported.");
      const normalizedMerges = [];
      for (const merge of sheet.merges) {
        const mergeRangeText = CE(normalizeWorksheetToolText(merge, 40)),
          rangeAddress = parseRangeAddress(mergeRangeText);
        ((rangeAddress.lastRow - rangeAddress.firstRow + 1) *
          (rangeAddress.lastCol - rangeAddress.firstCol + 1) >
          500 && failWithMessage("A merged area supports at most 500 cells."),
          normalizedMerges.some((existingMerge) =>
            doesRangeOverlap(existingMerge, mergeRangeText),
          ) && failWithMessage("Merged ranges cannot overlap."),
          normalizedMerges.push(mergeRangeText));
      }
      targetSheet.merges = normalizedMerges;
    }
    if (
      (sheet.validations !== undefined &&
        (targetSheet.validations = normalizeDataValidationRules(
          sheet.validations,
        )),
      sheet.conditionalFormats !== undefined &&
        (targetSheet.conditionalFormats = normalizeConditionalRules(
          sheet.conditionalFormats,
        )),
      sheet.protection !== undefined)
    ) {
      const protection = sheet.protection;
      ((!isNonArrayObject(protection) ||
        typeof protection.enabled != "boolean" ||
        (protection.hash !== undefined &&
          (!/^[a-f\d]{64}$/i.test(String(protection.hash)) ||
            typeof protection.salt != "string" ||
            protection.salt.length > 100))) &&
        failWithMessage("Invalid sheet protection."),
        (targetSheet.protection = {
          enabled: protection.enabled,
          ...(protection.hash
            ? { hash: String(protection.hash), salt: String(protection.salt) }
            : {}),
        }));
    }
  }
}
function mergeWorksheetCells(workbook, sheetId, rangeText, unmerge = false) {
  const sheet = workbook.sheets.find((sheet) => sheet.id === sheetId);
  assertSheetUnprotected(sheet);
  const normalizedRange = CE(rangeText),
    rangeAddress = parseRangeAddress(normalizedRange);
  (rangeAddress.lastRow - rangeAddress.firstRow + 1) *
    (rangeAddress.lastCol - rangeAddress.firstCol + 1) >
    500 && failWithMessage("Merge at most 500 cells.");
  const overlappingMerges = (sheet.merges ?? []).filter((existingMerge) =>
    doesRangeOverlap(existingMerge, normalizedRange),
  );
  !unmerge &&
    overlappingMerges.length &&
    failWithMessage("Unmerge overlapping areas first.");
  const cells = { ...sheet.cells };
  if (!unmerge) {
    const anchorAddress = `${columnIndexToLetters(rangeAddress.firstCol)}${rangeAddress.firstRow + 1}`;
    cells[anchorAddress] = {
      ...cells[anchorAddress],
      raw: cells[anchorAddress]?.raw ?? null,
      style: { ...cells[anchorAddress]?.style, align: "center" },
    };
  }
  return {
    ...workbook,
    sheets: workbook.sheets.map((sheet) =>
      sheet.id === sheetId
        ? {
            ...sheet,
            cells: cells,
            merges: unmerge
              ? (sheet.merges ?? []).filter(
                  (merge) => !overlappingMerges.includes(merge),
                )
              : [...(sheet.merges ?? []), normalizedRange],
          }
        : sheet,
    ),
  };
}
function oP(names) {
  let totalCellCount = 0;
  for (const { refersTo } of names) {
    const referenceText = stripFormulaStringLiterals(
      refersTo.replace(/"(?:[^"]|"")*"/g, '""').replace(/'(?:[^']|'')+'!/g, ""),
    );
    for (const match of referenceText.matchAll(
      /(?:^|[^A-Za-z0-9_.])((?:\$?[A-Za-z]{1,3}\$?[1-9]\d*)(?::\$?[A-Za-z]{1,3}\$?[1-9]\d*)?)(?![A-Za-z0-9_.])/g,
    )) {
      const rangeAddress = parseRangeAddress(match[1]);
      if (
        ((totalCellCount +=
          (rangeAddress.lastRow - rangeAddress.firstRow + 1) *
          (rangeAddress.lastCol - rangeAddress.firstCol + 1)),
        totalCellCount > workbookLimits.maxCells)
      )
        throw Error(
          "Named ranges exceed the 100,000-cell XLSX import/export processing budget. Use native JSON for larger definitions.",
        );
    }
  }
}
function aP(workbookXml) {
  const definedNamesXml = parseDefinedNamesXml(workbookXml),
    result = { names: [], warnings: [], builtinXml: "" };
  if (!definedNamesXml) return result;
  let body = definedNamesXml.body,
    definedNameCount = 0;
  const seenScopedNames = new Set(),
    sheetCount = [
      ...workbookXml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?sheet\b[^>]*>/g),
    ].length;
  for (; body.trim();) {
    const definedNameMatch =
      /^\s*<((?:[A-Za-z_][\w.-]*:)?definedName)\b((?:[^"'<>]|"[^"]*"|'[^']*')*?)(?:\/>|>([^<]*)<\/\1\s*>)/.exec(
        body,
      );
    if (!definedNameMatch)
      return throwXlsxDefinedNameError(
        "Malformed definedName content; only text definitions are supported.",
      );
    const definedNameXmlAttributes = parseDefinedNameXmlAttributes(
        definedNameMatch[2],
      ),
      nameAttribute = definedNameXmlAttributes.name ?? "",
      localSheetId = definedNameXmlAttributes.localSheetId,
      namespacePrefix = definedNameMatch[1].slice(0, -11);
    if (
      { ...definedNamesXml.namespaces, ...definedNameXmlAttributes }[
        namespacePrefix ? `xmlns:${namespacePrefix.slice(0, -1)}` : "xmlns"
      ] !== spreadsheetMlNamespace
    )
      return throwXlsxDefinedNameError(
        "Unsupported or undeclared definedName namespace.",
      );
    for (const definedNameXmlAttribute of Object.keys(definedNameXmlAttributes))
      if (
        definedNameXmlAttribute.startsWith("xmlns") &&
        definedNameXmlAttributes[definedNameXmlAttribute] !==
          spreadsheetMlNamespace
      )
        return throwXlsxDefinedNameError("Unsupported definedName namespace.");
    if (
      localSheetId !== undefined &&
      (!/^\d+$/.test(localSheetId) ||
        !Number.isSafeInteger(Number(localSheetId)))
    )
      return throwXlsxDefinedNameError(
        `Name "${nameAttribute}" has an invalid localSheetId.`,
      );
    const refersTo = `=${nP(
      (definedNameMatch[3] ?? "").replace(
        /\r\n?/g,
        `
`,
      ),
    ).replace(/^=/, "")}`;
    if (
      (validateNameFormulaText(refersTo),
      localSheetId !== undefined && Number(localSheetId) >= sheetCount)
    )
      return throwXlsxDefinedNameError(
        `Name "${nameAttribute}" has an out-of-range localSheetId.`,
      );
    const scopedNameKey = `${localSheetId === undefined ? "workbook" : Number(localSheetId)}:${nameAttribute.toLowerCase()}`;
    if (seenScopedNames.has(scopedNameKey))
      return throwXlsxDefinedNameError(
        `Duplicate name "${nameAttribute}" in the same scope.`,
      );
    if (
      (seenScopedNames.add(scopedNameKey),
      reservedDefinedNames.has(nameAttribute.toLowerCase()))
    )
      result.builtinXml += `<definedName${Object.entries(
        definedNameXmlAttributes,
      )
        .filter(([attributeKey]) => !attributeKey.startsWith("xmlns"))
        .map(
          ([attributeName, attributeValue]) =>
            ` ${attributeName}="${validateXmlSafeName(attributeValue, true)}"`,
        )
        .join("")}>${validateXmlSafeName(refersTo.slice(1))}</definedName>`;
    else if (nameAttribute.toLowerCase().startsWith("_xlnm.")) {
      if (++definedNameCount > 1e3)
        return throwXlsxDefinedNameError(
          "The workbook exceeds the 1,000-defined-name limit.",
        );
      result.warnings.push(
        `Built-in defined name "${nameAttribute}" is not supported and was omitted. Review its source feature before relying on imported formulas.`,
      );
    } else {
      if (++definedNameCount > 1e3)
        return throwXlsxDefinedNameError(
          "The workbook exceeds the 1,000-defined-name limit.",
        );
      for (const executableAttribute of ["function", "vbProcedure", "xlm"])
        if (
          definedNameXmlAttributes[executableAttribute] !== undefined &&
          !["0", "false"].includes(
            definedNameXmlAttributes[executableAttribute],
          )
        )
          return throwXlsxDefinedNameError(
            `Name "${nameAttribute}" has an executable ${executableAttribute} attribute.`,
          );
      const unsupportedAttributes = Object.keys(
        definedNameXmlAttributes,
      ).filter(
        (definedNameXmlAttribute) =>
          !["name", "localSheetId", "comment"].includes(
            definedNameXmlAttribute,
          ) && !definedNameXmlAttribute.startsWith("xmlns"),
      );
      (unsupportedAttributes.length &&
        result.warnings.push(
          `Defined name "${nameAttribute}" retained its definition and scope; unsupported attributes (${unsupportedAttributes.join(", ")}) were omitted. Hidden names are visible in Name Manager.`,
        ),
        result.names.push({
          name: nameAttribute,
          refersTo: refersTo,
          ...(localSheetId !== undefined
            ? { localSheetId: Number(localSheetId) }
            : {}),
          ...(definedNameXmlAttributes.comment !== undefined
            ? { comment: definedNameXmlAttributes.comment }
            : {}),
        }));
    }
    body = body.slice(definedNameMatch[0].length);
  }
  oP(result.names);
  try {
    validateDefinedName(
      result.names.map((definedName) => ({
        name: definedName.name,
        refersTo: definedName.refersTo,
        ...(definedName.localSheetId !== undefined
          ? { scopeSheetId: String(definedName.localSheetId) }
          : {}),
        ...(definedName.comment !== undefined
          ? { comment: definedName.comment }
          : {}),
      })),
      {
        sheets: [
          ...new Set(
            result.names.flatMap((nameEntry) =>
              nameEntry.localSheetId === undefined
                ? []
                : [nameEntry.localSheetId],
            ),
          ),
        ].map((sheetScopeId) => ({ id: String(sheetScopeId) })),
      },
    );
  } catch (error) {
    return throwXlsxDefinedNameError(error.message);
  }
  return result;
}

export {
  doesRangeOverlap,
  setOutlineGroupCollapsed,
  CE,
  removeOutlineGroup,
  findDataValidationsInRange,
  resolveFilteredRowContext,
  normalizeReviewMetadata,
  EQ,
  normalizeConditionalRules,
  parsePrintTitles,
  evaluateConditionalFormatRules,
  formatAxisIntervalRef,
  isCanonicalCellAddress,
  setFreezePanes,
  setSplitPanes,
  validateFreezePanes,
  validateWhatIfDefinition,
  enumerateSelectionCells,
  attachReviewThreadsToWorkbook,
  createAnchoredDrawingObject,
  clearAxisOutlineGroups,
  validateFormulaWatches,
  validateWorksheetFeatures,
  isCellAddressInRanges,
  validateHyperlinkRecord,
  applyChartDrawingLayout,
  normalizeDataValidationRules,
  _t,
  aP,
  parseRangeAddress,
  setHiddenAxisRange,
  shiftFormulaWatchesOnAxisEdit,
  buildDrawingExportNotes,
  parseSelectionAreas,
  shiftPrintSettingsOnAxisEdit,
  mergeWorksheetCells,
  setOutlineGroupLevel,
  lD,
  oP,
  buildBorderPresetStyles,
  pQ,
  normalizeCellAddress,
  shiftReviewTargetsOnAxisEdit,
  validateSavedWhatIfAnalyses,
  analysisOutputEvidenceSchema,
  setFreezePanesAtCell,
  applyAxisStructureEdit,
  countSelectionCells,
  yQ,
};

/* Also used by the workers. */
export {
  evidenceAddressSchema,
  evidenceCellRecordSchema,
  evidenceFieldValueSchema,
  evidenceSourceOriginSchema,
  evidenceTransformStepSchema,
};

/* Also used by the workers. */
export {
  normalizeFormulaWatchAddress,
  normalizeWhatIfCell,
  normalizeWhatIfCellReferences,
  normalizeWorksheetToolText,
};
