import { throwErrorMessage, failWithMessage } from "./03-workbook-model.js";
import { formatCellAddress } from "./04-cell-input-paste-and-fill.js";
import {
  buildCompositeKey,
  isFormulaInputValue,
} from "./11-local-workbook-store-and-versions.js";
import { X } from "lucide-react";
import i from "react/jsx-runtime";
import ne from "react";
import { workbookLimits } from "./02-text-and-financial-functions.js";
import {
  calculateWorkbookWithSettings,
  cn,
  parseRangeAddress,
  resolveDefinedNameValue,
  columnIndexToLetters,
  matrixResultFunctionNames,
} from "./03-workbook-model.js";
import {
  useWorkbookMessages,
  LocalizedText,
  LocalizedNumberField,
} from "./04-cell-input-paste-and-fill.js";
import {
  importNativeJsonFile,
  toFileProcessingError,
} from "./07-file-import-and-worker-dispatch.js";
import { useCancellableOperation } from "./09-workbook-links.js";
import { inspectFormulaDependencies } from "./12-consolidation-and-dependency-inspection.js";
import {
  parseAssistantChatRequest,
  validateAssistantProposal,
  LLMRequestError,
  raceWithAbortSignal,
  validateCalculationClock,
} from "./20-calculation-options-and-assistant.js";

async function requestAssistantAnswer(
  adapter,
  prompt,
  assistantSelectionContext,
  options = {},
) {
  if (options.reliability) {
    const { reliability, ...remainingOptions } = options;
    return reliability.run("provider", "request", () =>
      requestAssistantAnswer(
        adapter,
        prompt,
        assistantSelectionContext,
        remainingOptions,
      ),
    );
  }
  if (
    typeof prompt != "string" ||
    !prompt.trim() ||
    prompt.length > 1e4 ||
    !Array.isArray(assistantSelectionContext?.cells) ||
    assistantSelectionContext.cells.length > workbookLimits.maxLLMCells
  )
    throw new LLMRequestError("invalid_request");
  validateCalculationClock(assistantSelectionContext);
  const selectionContextCopy = structuredClone(assistantSelectionContext),
    chatRequest = options.conversation
      ? parseAssistantChatRequest(options.conversation)
      : undefined;
  if (
    (chatRequest?.attachments.length ||
      chatRequest?.messages.some((message) => message.attachments?.length)) &&
    !adapter.chat?.attachments
  )
    throw new Error(
      "This provider does not support file attachments. Remove the files or use a chat-enabled host adapter.",
    );
  if (chatRequest?.messages.length && !adapter.chat?.history)
    throw new Error(
      "This provider does not support conversation history. Start a new chat or use a chat-enabled host adapter.",
    );
  const proposal = await raceWithAbortSignal(
    (abortSignal) =>
      chatRequest
        ? adapter.propose(
            prompt,
            structuredClone(selectionContextCopy),
            abortSignal,
            structuredClone(chatRequest),
          )
        : adapter.propose(
            prompt,
            structuredClone(selectionContextCopy),
            abortSignal,
          ),
    options.signal,
    options.timeoutMs,
  );
  try {
    return validateAssistantProposal(proposal, selectionContextCopy);
  } catch {
    throw new LLMRequestError("malformed_response");
  }
}
const materialityScreenLimits = { sourceCells: 500, pageSize: 20 },
  Tz = new WeakMap();

const cde = 864e5;
function parseStrictIsoDate(dateText) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(dateText)) return null;
  const timestamp = Date.parse(`${dateText}T00:00:00.000Z`);
  return Number.isFinite(timestamp) &&
    new Date(timestamp).toISOString().slice(0, 10) === dateText
    ? timestamp
    : null;
}
function normalizeDateToIso(dateValue) {
  return typeof dateValue == "string"
    ? parseStrictIsoDate(dateValue) === null
      ? null
      : dateValue
    : typeof dateValue != "number" ||
        !Number.isInteger(dateValue) ||
        dateValue < 1 ||
        dateValue > 2958465 ||
        dateValue === 60
      ? null
      : new Date(
          Date.UTC(1899, 11, 31) +
            (dateValue > 60 ? dateValue - 1 : dateValue) * cde,
        )
          .toISOString()
          .slice(0, 10);
}
function normalizeScreenHostScope(workbook, scopeOptions) {
  return (
    scopeOptions.includeHidden !== undefined &&
      typeof scopeOptions.includeHidden != "boolean" &&
      failWithMessage("Include hidden must be boolean."),
    scopeOptions.readOnly !== undefined &&
      typeof scopeOptions.readOnly != "boolean" &&
      failWithMessage("Read-only must be boolean."),
    scopeOptions.allowedSheetIds !== undefined &&
      (!Array.isArray(scopeOptions.allowedSheetIds) ||
        scopeOptions.allowedSheetIds.some(
          (allowedSheetId) =>
            typeof allowedSheetId != "string" ||
            !workbook.sheets.some((sheet) => sheet.id === allowedSheetId),
        )) &&
      failWithMessage("The materiality screen host context is invalid."),
    {
      includeHidden: !!scopeOptions.includeHidden,
      allowedSheetIds: [
        ...new Set(
          scopeOptions.allowedSheetIds ??
            workbook.sheets.map((sheet) => sheet.id),
        ),
      ].sort(),
    }
  );
}
function normalizeMaterialityConfig(workbook, config) {
  const allowedKeys = [
    "sheetId",
    "range",
    "headers",
    "amountColumn",
    "currencyColumn",
    "dateColumn",
    "currency",
    "threshold",
    "amountMode",
    "comparator",
    "effectiveFrom",
    "effectiveTo",
  ];
  ((!config ||
    typeof config != "object" ||
    Object.keys(config).some(
      (unknownKey) => !allowedKeys.includes(unknownKey),
    ) ||
    allowedKeys.some((missingKey) => !(missingKey in config))) &&
    failWithMessage(
      "Provide a complete, supported materiality screen configuration.",
    ),
    (typeof config.sheetId != "string" ||
      !workbook.sheets.some((sheet) => sheet.id === config.sheetId)) &&
      failWithMessage("The source worksheet no longer exists."),
    (typeof config.range != "string" ||
      !/^\$?[A-Za-z]{1,3}\$?[1-9]\d*(?::\$?[A-Za-z]{1,3}\$?[1-9]\d*)?$/.test(
        config.range,
      )) &&
      failWithMessage("Choose one rectangular A1 source range."));
  const rangeAddress = parseRangeAddress(config.range);
  ((rangeAddress.lastRow - rangeAddress.firstRow + 1) *
    (rangeAddress.lastCol - rangeAddress.firstCol + 1) >
    materialityScreenLimits.sourceCells &&
    failWithMessage(
      "The materiality screen supports at most 500 selected source cells.",
    ),
    (typeof config.headers != "boolean" ||
      rangeAddress.lastRow - rangeAddress.firstRow + 1 <=
        Number(config.headers)) &&
      failWithMessage(
        "Choose a source range with at least one data row and valid header handling.",
      ));
  const columnIndexes = [
    config.amountColumn,
    config.currencyColumn,
    config.dateColumn,
  ];
  return (
    (columnIndexes.some(
      (columnIndex) =>
        !Number.isInteger(columnIndex) ||
        columnIndex < rangeAddress.firstCol ||
        columnIndex > rangeAddress.lastCol,
    ) ||
      new Set(columnIndexes).size !== 3) &&
      failWithMessage(
        "Choose three distinct amount, currency and date columns inside the source range.",
      ),
    (typeof config.threshold != "number" ||
      !Number.isFinite(config.threshold) ||
      config.threshold < 0) &&
      failWithMessage("Threshold must be a finite, nonnegative number."),
    (!["absolute", "signed"].includes(config.amountMode) ||
      !["gt", "gte", "lt", "lte", "eq"].includes(config.comparator)) &&
      failWithMessage(
        "Choose an absolute or signed amount and a supported comparison rule.",
      ),
    (typeof config.currency != "string" ||
      !/^[A-Za-z]{3}$/.test(config.currency)) &&
      failWithMessage(
        "Declare a three-letter currency identifier; no currency conversion is performed.",
      ),
    (typeof config.effectiveFrom != "string" ||
      typeof config.effectiveTo != "string" ||
      parseStrictIsoDate(config.effectiveFrom) === null ||
      parseStrictIsoDate(config.effectiveTo) === null ||
      config.effectiveFrom > config.effectiveTo) &&
      failWithMessage(
        "Choose valid ISO start and end dates in chronological order.",
      ),
    {
      sheetId: config.sheetId,
      range: `${formatCellAddress(rangeAddress.firstRow, rangeAddress.firstCol)}:${formatCellAddress(rangeAddress.lastRow, rangeAddress.lastCol)}`,
      headers: config.headers,
      amountColumn: config.amountColumn,
      currencyColumn: config.currencyColumn,
      dateColumn: config.dateColumn,
      currency: config.currency.toUpperCase(),
      threshold: config.threshold,
      amountMode: config.amountMode,
      comparator: config.comparator,
      effectiveFrom: config.effectiveFrom,
      effectiveTo: config.effectiveTo,
    }
  );
}
function resolveMaterialitySourceRange(workbook, config, screenHostScope) {
  const sheet = workbook.sheets.find((sheet) => sheet.id === config.sheetId);
  (!sheet || !screenHostScope.allowedSheetIds.includes(sheet.id)) &&
    failWithMessage("The source is outside the authorized host context.");
  const rangeAddress = parseRangeAddress(config.range);
  return (
    !screenHostScope.includeHidden &&
      (sheet.hidden ||
        sheet.veryHidden ||
        cn(sheet, "row").some(
          (rowSpan) =>
            rowSpan.start <= rangeAddress.lastRow &&
            rowSpan.end >= rangeAddress.firstRow,
        ) ||
        cn(sheet, "column").some(
          (columnSpan) =>
            columnSpan.start <= rangeAddress.lastCol &&
            columnSpan.end >= rangeAddress.firstCol,
        )) &&
      failWithMessage(
        "The source range includes hidden content. Explicitly include hidden content only within authorized host scope.",
      ),
    (sheet.merges ?? []).some((mergeRange) => {
      const mergeAddress = parseRangeAddress(mergeRange);
      return (
        mergeAddress.firstRow <= rangeAddress.lastRow &&
        mergeAddress.lastRow >= rangeAddress.firstRow &&
        mergeAddress.firstCol <= rangeAddress.lastCol &&
        mergeAddress.lastCol >= rangeAddress.firstCol
      );
    }) &&
      failWithMessage(
        "Unmerge the source range before screening so every record has unambiguous fields.",
      ),
    sheet
  );
}
function inspectScreenFormulaCells(
  workbook,
  materialitySourceRange,
  cellAddresses,
  screenHostScope,
) {
  const formulaAddresses = cellAddresses.filter((candidateAddress) =>
      isFormulaInputValue(materialitySourceRange.cells[candidateAddress]),
    ),
    rawValueByAddress = new Map(
      cellAddresses.map((address) => [
        address,
        materialitySourceRange.cells[address]?.raw ?? null,
      ]),
    );
  if (!formulaAddresses.length) return rawValueByAddress;
  const workbookJson = JSON.stringify(workbook);
  let isolatedSheetId = "";
  for (
    let attempt = 0;
    attempt < 4 &&
    ((isolatedSheetId = `_ms_${crypto.randomUUID().replace(/-/g, "").slice(0, 24)}`),
    !!workbookJson.toLowerCase().includes(isolatedSheetId.toLowerCase()));
    attempt++
  )
    isolatedSheetId = "";
  isolatedSheetId ||
    failWithMessage("A private calculation context could not be isolated.");
  const inspectionSheet = {
    id: isolatedSheetId,
    name: isolatedSheetId,
    rowCount: workbookLimits.maxRows,
    colCount: workbookLimits.maxCols,
    cells: {
      A1: {
        raw: `=SUM(${formulaAddresses.map((formulaAddress) => `'${materialitySourceRange.name.replace(/'/g, "''")}'!${formulaAddress}`).join(",")})`,
      },
    },
  };
  String(inspectionSheet.cells.A1.raw).length > 8192 &&
    failWithMessage(
      "The selected formulas exceed the bounded calculation inspection length.",
    );
  const dependencyGraph = inspectFormulaDependencies(
      {
        ...workbook,
        sheets: [
          inspectionSheet,
          ...workbook.sheets.map((sheet) =>
            sheet.veryHidden ? { ...sheet, hidden: true } : sheet,
          ),
        ],
      },
      {
        sheetId: inspectionSheet.id,
        address: "A1",
        direction: "precedents",
        depth: 8,
        includeHidden: screenHostScope.includeHidden,
        allowedSheetIds: [
          ...screenHostScope.allowedSheetIds,
          inspectionSheet.id,
        ],
        limits: { maxNodes: 1e3, maxFormulaCells: 5e3, maxReferences: 1e4 },
      },
    ),
    closureKeys = new Set(dependencyGraph.nodes.map((node) => node.key));
  (dependencyGraph.coverage.truncated ||
    dependencyGraph.coverage.depthLimited ||
    dependencyGraph.cycles.length ||
    dependencyGraph.issues.some(
      (issue) =>
        issue.from === dependencyGraph.target.key ||
        closureKeys.has(issue.from),
    )) &&
    failWithMessage(
      "Formula fields require a complete supported static dependency closure within the visible authorized context (1,000 cells, depth 8, 5,000 scanned formulas).",
    );
  const volatileFunctionNames = new Set([
    ...matrixResultFunctionNames,
    "ISFORMULA",
    "INDEX",
    "CHOOSE",
    "INDIRECT",
    "OFFSET",
    "LET",
    "LAMBDA",
    "RAND",
    "RANDBETWEEN",
    "NOW",
    "TODAY",
    "CELL",
    "INFO",
    "RTD",
    "WEBSERVICE",
    "STOCKHISTORY",
    "EVALUATE",
    "GET.CELL",
  ]);
  for (const node of dependencyGraph.nodes) {
    const cell = workbook.sheets.find((sheet) => sheet.id === node.sheetId)
      ?.cells[node.address];
    if (
      ((!cell || cell.raw === null) &&
        failWithMessage(
          "Formula dependencies must be populated scalar cells; ambiguous blank or spill-child dependencies cannot be screened.",
        ),
      isFormulaInputValue(cell))
    ) {
      const resolvedFormula = resolveDefinedNameValue(
        String(cell.raw),
        workbook,
        node.sheetId,
      ).replace(/"(?:[^"]|"")*"|'(?:[^']|'')*'(?=\s*!)/g, "");
      (/[{}#]/.test(resolvedFormula) ||
        [...resolvedFormula.matchAll(/([A-Za-z_][A-Za-z0-9_.]*)\s*\(/g)].some(
          (functionMatch) =>
            volatileFunctionNames.has(
              functionMatch[1]
                .toUpperCase()
                .replace(/^_XLFN\./, "")
                .replace(/^_XLWS\./, ""),
            ),
        )) &&
        failWithMessage(
          "The screen excludes volatile, dynamic-reference and array/spill formulas. Use supported scalar values or formulas.",
        );
    }
  }
  const scopedWorkbook = {
      ...workbook,
      sheets: workbook.sheets
        .filter((sheet) => screenHostScope.allowedSheetIds.includes(sheet.id))
        .map((scopedSheet) => ({
          ...scopedSheet,
          cells: Object.fromEntries(
            Object.entries(scopedSheet.cells).filter(([cellAddress]) =>
              closureKeys.has(buildCompositeKey(scopedSheet.id, cellAddress)),
            ),
          ),
        })),
    },
    workbookWithSettings = calculateWorkbookWithSettings(scopedWorkbook);
  for (const node of dependencyGraph.nodes) {
    const computedValue = workbookWithSettings(node.sheetId, node.address);
    (workbookWithSettings.getOperationalErrors().length &&
      failWithMessage(
        "An operational calculation limit or cycle prevents screening, even when a formula handles the visible error.",
      ),
      workbookWithSettings.getSpillRange(node.sheetId, node.address) &&
        failWithMessage(
          "The screen requires scalar fields; a formula produced an array/spill result.",
        ),
      isFormulaInputValue(
        workbook.sheets.find((sheet) => sheet.id === node.sheetId)?.cells[
          node.address
        ],
      ) &&
        typeof computedValue == "string" &&
        /^#(?:[A-Z0-9_/.]+[!?]?)/.test(computedValue) &&
        failWithMessage(
          "A selected formula or dependency has a calculation error. Correct it before screening.",
        ),
      node.sheetId === materialitySourceRange.id &&
        formulaAddresses.includes(node.address) &&
        rawValueByAddress.set(node.address, computedValue));
  }
  return rawValueByAddress;
}
const compareAmountToThreshold = (amount, threshold, comparator) =>
  comparator === "gt"
    ? amount > threshold
    : comparator === "gte"
      ? amount >= threshold
      : comparator === "lt"
        ? amount < threshold
        : comparator === "lte"
          ? amount <= threshold
          : amount === threshold;
function normalizeMaterialityRequest(workbook, config, scopeOptions = {}) {
  const materialityConfig = normalizeMaterialityConfig(workbook, config);
  return {
    ...buildMaterialitySourceEvidence(
      workbook,
      {
        sheetId: materialityConfig.sheetId,
        range: materialityConfig.range,
        headers: materialityConfig.headers,
        columns: [
          materialityConfig.amountColumn,
          materialityConfig.currencyColumn,
          materialityConfig.dateColumn,
        ],
      },
      scopeOptions,
    ),
    request: materialityConfig,
  };
}
function buildMaterialitySourceEvidence(
  workbook,
  sourceSelection,
  scopeOptions = {},
) {
  (!sourceSelection ||
    typeof sourceSelection.sheetId != "string" ||
    typeof sourceSelection.range != "string" ||
    typeof sourceSelection.headers != "boolean" ||
    !Array.isArray(sourceSelection.columns) ||
    sourceSelection.columns.length > 500) &&
    failWithMessage(
      "Provide a bounded source range and scalar field selection.",
    );
  const rangeAddress = parseRangeAddress(sourceSelection.range),
    sheet = workbook.sheets.find(
      (sheet) => sheet.id === sourceSelection.sheetId,
    );
  (!sheet ||
    rangeAddress.lastRow >= sheet.rowCount ||
    rangeAddress.lastCol >= sheet.colCount ||
    (rangeAddress.lastRow - rangeAddress.firstRow + 1) *
      (rangeAddress.lastCol - rangeAddress.firstCol + 1) >
      materialityScreenLimits.sourceCells ||
    sourceSelection.columns.some(
      (column) =>
        !Number.isInteger(column) ||
        column < rangeAddress.firstCol ||
        column > rangeAddress.lastCol,
    )) &&
    failWithMessage(
      "Choose scalar fields inside one source rectangle of at most 500 cells and within worksheet dimensions.",
    );
  const normalizedSelection = {
      ...sourceSelection,
      range: `${formatCellAddress(rangeAddress.firstRow, rangeAddress.firstCol)}:${formatCellAddress(rangeAddress.lastRow, rangeAddress.lastCol)}`,
    },
    screenHostScope = normalizeScreenHostScope(workbook, scopeOptions),
    materialitySourceRange = resolveMaterialitySourceRange(
      workbook,
      normalizedSelection,
      screenHostScope,
    ),
    u = parseRangeAddress(normalizedSelection.range);
  for (let firstRow = u.firstRow; firstRow <= u.lastRow; firstRow++)
    for (let firstCol = u.firstCol; firstCol <= u.lastCol; firstCol++) {
      const rawValue =
        materialitySourceRange.cells[formatCellAddress(firstRow, firstCol)]
          ?.raw;
      typeof rawValue == "number" &&
        !Number.isFinite(rawValue) &&
        failWithMessage(
          "Selected raw evidence contains a nonfinite number and cannot be preserved in JSON. Correct it before screening.",
        );
    }
  const inspectedAddresses = [];
  for (
    let rowIndex = u.firstRow + Number(normalizedSelection.headers);
    rowIndex <= u.lastRow;
    rowIndex++
  )
    for (const columnIndex of new Set(normalizedSelection.columns))
      inspectedAddresses.push(formatCellAddress(rowIndex, columnIndex));
  const formulaValues = inspectScreenFormulaCells(
      workbook,
      materialitySourceRange,
      inspectedAddresses,
      screenHostScope,
    ),
    rows = [];
  for (
    let dataRowIndex = u.firstRow + Number(normalizedSelection.headers);
    dataRowIndex <= u.lastRow;
    dataRowIndex++
  )
    rows.push({
      sourceRow: dataRowIndex + 1,
      sourceRange: `${formatCellAddress(dataRowIndex, u.firstCol)}:${formatCellAddress(dataRowIndex, u.lastCol)}`,
      cells: Array.from(
        { length: u.lastCol - u.firstCol + 1 },
        (v, columnOffset) =>
          structuredClone(
            materialitySourceRange.cells[
              formatCellAddress(dataRowIndex, u.firstCol + columnOffset)
            ] ?? { raw: null },
          ),
      ),
      values: Array.from(
        { length: u.lastCol - u.firstCol + 1 },
        (v, valueOffset) => {
          const cellAddress = formatCellAddress(
            dataRowIndex,
            u.firstCol + valueOffset,
          );
          return formulaValues.has(cellAddress)
            ? formulaValues.get(cellAddress)
            : (materialitySourceRange.cells[cellAddress]?.raw ?? null);
        },
      ),
    });
  const evidence = {
    context: screenHostScope,
    sheetId: materialitySourceRange.id,
    sheetName: materialitySourceRange.name,
    rows: rows,
    headers: Array.from(
      { length: u.lastCol - u.firstCol + 1 },
      (m, headerOffset) =>
        normalizedSelection.headers
          ? String(
              materialitySourceRange.cells[
                formatCellAddress(u.firstRow, u.firstCol + headerOffset)
              ]?.raw ??
                `Column ${columnIndexToLetters(u.firstCol + headerOffset)}`,
            )
          : `Column ${columnIndexToLetters(u.firstCol + headerOffset)}`,
    ),
  };
  return (
    new TextEncoder().encode(JSON.stringify(evidence, null, 2)).length > 2e6 &&
      failWithMessage(
        "The complete source evidence exceeds 2,000,000 bytes. Select a smaller population.",
      ),
    evidence
  );
}
function runMaterialityScreen(workbook, screenConfig, scopeOptions = {}) {
  const materialityRequest = normalizeMaterialityRequest(
      workbook,
      screenConfig,
      scopeOptions,
    ),
    { request, context } = materialityRequest,
    rangeAddress = parseRangeAddress(request.range),
    rows = [],
    counts = {
      population: 0,
      eligible: 0,
      screened: 0,
      flagged: 0,
      notFlagged: 0,
      excluded: 0,
      notScreened: 0,
      currencyExceptions: 0,
      invalidAmounts: 0,
      invalidDates: 0,
      outsidePeriod: 0,
    };
  let knownSourceAmount = 0,
    eligibleAmount = 0,
    unknownOrOtherCurrencyRows = 0;
  for (const row of materialityRequest.rows) {
    const valueAt = (columnIndex) =>
        row.values[columnIndex - rangeAddress.firstCol],
      rawAmount = valueAt(request.amountColumn),
      rawCurrency = valueAt(request.currencyColumn),
      rawDate = valueAt(request.dateColumn),
      amount =
        typeof rawAmount == "number" && Number.isFinite(rawAmount)
          ? rawAmount
          : null,
      currency =
        typeof rawCurrency == "string" && /^[A-Za-z]{3}$/.test(rawCurrency)
          ? rawCurrency.toUpperCase()
          : null,
      dateToIso = normalizeDateToIso(rawDate),
      reasons = [];
    (amount === null &&
      (reasons.push(
        "Amount is missing or is not a finite numeric value (numeric text is not coerced).",
      ),
      counts.invalidAmounts++),
      (!currency || currency !== request.currency) &&
        (reasons.push(
          currency
            ? `Currency ${currency} differs from the declared ${request.currency}; conversion is unsupported.`
            : "Currency is missing or is not a three-letter identifier.",
        ),
        counts.currencyExceptions++),
      dateToIso
        ? (dateToIso < request.effectiveFrom ||
            dateToIso > request.effectiveTo) &&
          (reasons.push("Date is outside the inclusive effective period."),
          counts.outsidePeriod++)
        : (reasons.push(
            "Date is missing or invalid; use ISO YYYY-MM-DD or an integer Excel 1900-system date serial (except 60).",
          ),
          counts.invalidDates++));
    const comparedAmount =
      amount === null
        ? null
        : request.amountMode === "absolute"
          ? Math.abs(amount)
          : amount;
    (amount !== null && currency === request.currency
      ? (knownSourceAmount += amount)
      : unknownOrOtherCurrencyRows++,
      reasons.length || (counts.eligible++, (eligibleAmount += amount)),
      rows.push({
        ...row,
        amount: amount,
        currency: currency,
        date: dateToIso,
        comparedAmount: comparedAmount,
        disposition: reasons.length
          ? "excluded"
          : compareAmountToThreshold(
                comparedAmount,
                request.threshold,
                request.comparator,
              )
            ? "flagged"
            : "not-flagged",
        reasons: reasons,
      }));
  }
  if (counts.currencyExceptions)
    for (const row of rows)
      row.disposition !== "excluded" &&
        ((row.disposition = "not-screened"),
        row.reasons.push(
          "Screen blocked by currency exceptions; no pass or exception conclusion is available.",
        ));
  counts.population = rows.length;
  for (const row of rows)
    row.disposition === "flagged"
      ? counts.flagged++
      : row.disposition === "not-flagged"
        ? counts.notFlagged++
        : row.disposition === "excluded"
          ? counts.excluded++
          : counts.notScreened++;
  counts.screened = counts.flagged + counts.notFlagged;
  const flaggedAmount = counts.currencyExceptions
    ? null
    : rows
        .filter((row) => row.disposition === "flagged")
        .reduce(
          (flaggedTotal, flaggedRow) => flaggedTotal + flaggedRow.amount,
          0,
        );
  [
    knownSourceAmount,
    eligibleAmount,
    knownSourceAmount - eligibleAmount,
    flaggedAmount,
  ].some(
    (controlTotal) => controlTotal !== null && !Number.isFinite(controlTotal),
  ) &&
    failWithMessage(
      "The control total exceeds finite numeric capacity. Reduce the population or magnitude.",
    );
  const screenReport = {
    schemaVersion: 1,
    kind: "materiality-screen",
    algorithm: "ledgerglass-materiality-1",
    workbookId: workbook.id,
    revision: workbook.revision,
    sheetId: materialityRequest.sheetId,
    sheetName: materialityRequest.sheetName,
    generatedAt: new Date().toISOString(),
    request: structuredClone(request),
    context,
    status: counts.currencyExceptions
      ? "blocked"
      : counts.excluded
        ? "complete-with-exclusions"
        : "complete",
    headers: materialityRequest.headers,
    rows: rows,
    counts: counts,
    totals: {
      currency: request.currency,
      knownSourceAmount: knownSourceAmount,
      eligibleAmount: eligibleAmount,
      excludedKnownAmount: knownSourceAmount - eligibleAmount,
      flaggedAmount: flaggedAmount,
      unknownOrOtherCurrencyRows: unknownOrOtherCurrencyRows,
    },
    notes: [
      "User-configured quantitative screen; it does not determine appropriate audit materiality or identify fraud.",
      "Period endpoints are inclusive. Comparisons use exact stored/computed binary numbers with no display rounding, epsilon or currency conversion.",
      "Every data row is accounted for. Headers alone are excluded automatically; total and subtotal rows are ordinary records unless the selected population excludes them.",
      "Amount/currency/date formulas use a bounded supported static scalar closure. Other evidence columns retain stored values/formula text. Blank fields are reported as missing, never inferred from spill results.",
      "Control totals retain signed amounts and cover only valid numeric amounts in the declared currency; exclusions and unknown or other-currency rows are disclosed.",
      "This report is a local snapshot with source locators; it is not immutable evidence, a signature or a source-lineage guarantee.",
    ],
  };
  return (
    new TextEncoder().encode(JSON.stringify(screenReport, null, 2)).length >
      2e6 &&
      failWithMessage(
        "The complete evidence report exceeds 2,000,000 bytes. Select a smaller source population.",
      ),
    Tz.set(screenReport, {
      snapshot: JSON.stringify(workbook),
      request: JSON.stringify(request),
      context: JSON.stringify(context),
      content: JSON.stringify(screenReport),
    }),
    screenReport
  );
}
function checkMaterialityReportFreshness(
  workbook,
  report,
  configInput,
  scopeOptions = {},
) {
  const snapshotRecord = Tz.get(report);
  if (!snapshotRecord)
    return {
      current: false,
      disclose: false,
      reason: "This report is untrusted or changed. Run the screen again.",
    };
  try {
    if (snapshotRecord.content !== JSON.stringify(report))
      return {
        current: false,
        disclose: false,
        reason: "This report is untrusted or changed. Run the screen again.",
      };
    const screenHostScope = normalizeScreenHostScope(workbook, scopeOptions);
    if (
      (resolveMaterialitySourceRange(workbook, report.request, screenHostScope),
      JSON.stringify(screenHostScope) !== snapshotRecord.context)
    )
      return {
        current: false,
        disclose: false,
        reason:
          "The authorized or hidden-content scope changed. Run the screen again.",
      };
    if (JSON.stringify(workbook) !== snapshotRecord.snapshot)
      return {
        current: false,
        disclose: false,
        reason: "The source workbook changed. Run the screen again.",
      };
    const materialityConfig = normalizeMaterialityConfig(workbook, configInput);
    return JSON.stringify(materialityConfig) !== snapshotRecord.request
      ? {
          current: false,
          disclose: true,
          reason:
            "The screen parameters changed. This is the previous result; run the screen again.",
        }
      : {
          current: true,
          disclose: true,
          reason: "Current result for this source and configuration.",
        };
  } catch {
    return {
      current: false,
      disclose: false,
      reason:
        "The source, settings or authorized scope is no longer valid. Run the screen again.",
    };
  }
}
function exportMaterialityReportJson(
  workbook,
  report,
  configInput,
  scopeOptions = {},
) {
  const freshness = checkMaterialityReportFreshness(
    workbook,
    report,
    configInput,
    scopeOptions,
  );
  return (
    freshness.current || failWithMessage(freshness.reason),
    JSON.stringify(report, null, 2)
  );
}
function parseMaterialityReportText(workbook, reportText, scopeOptions = {}) {
  (typeof reportText != "string" ||
    reportText.length > 2e6 ||
    new TextEncoder().encode(reportText).length > 2e6) &&
    failWithMessage(
      "The screen configuration/report must be at most 2,000,000 bytes.",
    );
  let parsedReport;
  try {
    parsedReport = JSON.parse(reportText);
  } catch {
    failWithMessage("The screen configuration/report is not valid JSON.");
  }
  return loadMaterialityConfiguration(workbook, parsedReport, scopeOptions);
}
function loadMaterialityConfiguration(
  workbook,
  savedReport,
  scopeOptions = {},
) {
  (!savedReport ||
    typeof savedReport != "object" ||
    Array.isArray(savedReport)) &&
    failWithMessage("Choose a supported screen configuration or report.");
  const report = savedReport;
  (report.schemaVersion !== 1 || report.kind !== "materiality-screen") &&
    failWithMessage(
      "The materiality screen configuration version is unsupported.",
    );
  const materialityConfig = normalizeMaterialityConfig(
    workbook,
    report.request,
  );
  return (
    resolveMaterialitySourceRange(
      workbook,
      materialityConfig,
      normalizeScreenHostScope(workbook, scopeOptions),
    ),
    materialityConfig
  );
}
const agingScheduleLimits = {
    sourceCells: 500,
    buckets: 24,
    pageSize: 20,
    reportBytes: 2e6,
  },
  _z = new WeakMap(),
  dayInMilliseconds = 864e5;
function parseAgingIsoDate(dateText) {
  if (
    typeof dateText != "string" ||
    !/^\d{4}-\d{2}-\d{2}$/.test(dateText) ||
    dateText.startsWith("0000-")
  )
    return null;
  const timestamp = Date.parse(`${dateText}T00:00:00.000Z`);
  return Number.isFinite(timestamp) &&
    new Date(timestamp).toISOString().slice(0, 10) === dateText
    ? timestamp / dayInMilliseconds
    : null;
}
function normalizeAgingDateInput(dateValue, dateSystem) {
  if (typeof dateValue == "string")
    return parseAgingIsoDate(dateValue) === null ? null : dateValue;
  if (
    typeof dateValue != "number" ||
    !Number.isInteger(dateValue) ||
    (dateSystem === "1900" &&
      (dateValue < 1 || dateValue > 2958465 || dateValue === 60)) ||
    (dateSystem === "1904" && (dateValue < 0 || dateValue > 2957003))
  )
    return null;
  const epochStart =
    dateSystem === "1900" ? Date.UTC(1899, 11, 31) : Date.UTC(1904, 0, 1);
  return new Date(
    epochStart +
      (dateSystem === "1900" && dateValue > 60 ? dateValue - 1 : dateValue) *
        dayInMilliseconds,
  )
    .toISOString()
    .slice(0, 10);
}
function defaultMaterialityConfig(agingConfiguration) {
  return {
    sheetId: agingConfiguration.sheetId,
    range: agingConfiguration.range,
    headers: agingConfiguration.headers,
    amountColumn: agingConfiguration.amountColumn,
    currencyColumn: agingConfiguration.currencyColumn,
    dateColumn: agingConfiguration.dateColumn,
    currency: agingConfiguration.currency,
    threshold: 0,
    amountMode: "absolute",
    comparator: "gte",
    effectiveFrom: "0000-01-01",
    effectiveTo: "9999-12-31",
  };
}
function normalizeAgingConfiguration(
  workbook,
  agingConfiguration,
  scopeOptions,
) {
  const requiredKeys = [
    "sheetId",
    "range",
    "headers",
    "amountColumn",
    "currencyColumn",
    "dateColumn",
    "currency",
    "dateBasis",
    "dateSystem",
    "asOf",
    "buckets",
  ];
  ((!agingConfiguration ||
    typeof agingConfiguration != "object" ||
    Array.isArray(agingConfiguration) ||
    Object.keys(agingConfiguration).length !== requiredKeys.length ||
    requiredKeys.some(
      (requiredKey) => !Object.hasOwn(agingConfiguration, requiredKey),
    )) &&
    throwErrorMessage("Provide a complete, supported aging configuration."),
    (!["due", "document"].includes(agingConfiguration.dateBasis) ||
      !["1900", "1904"].includes(agingConfiguration.dateSystem)) &&
      throwErrorMessage(
        "Choose due or document date and an explicit 1900 or 1904 serial date system.",
      ),
    parseAgingIsoDate(agingConfiguration.asOf) === null &&
      throwErrorMessage(
        "Choose a valid fixed as-of date in ISO YYYY-MM-DD format.",
      ),
    (!Array.isArray(agingConfiguration.buckets) ||
      agingConfiguration.buckets.length < 1 ||
      agingConfiguration.buckets.length > agingScheduleLimits.buckets) &&
      throwErrorMessage("Configure between 1 and 24 aging buckets."));
  let expectedBucketStart = 0;
  const bucketLabels = new Set();
  for (const [bucketIndex, bucket] of agingConfiguration.buckets.entries())
    ((!bucket ||
      typeof bucket != "object" ||
      Array.isArray(bucket) ||
      Object.keys(bucket).length !== 3 ||
      !["label", "from", "to"].every((bucketKey) =>
        Object.hasOwn(bucket, bucketKey),
      ) ||
      typeof bucket.label != "string" ||
      !bucket.label.trim() ||
      bucket.label.trim().length > 80 ||
      /[\u0000-\u001f\u007f]/.test(bucket.label) ||
      bucketLabels.has(bucket.label.trim().toLowerCase())) &&
      throwErrorMessage(
        "Each bucket needs a unique label of 1–80 characters and inclusive from/to bounds.",
      ),
      bucketLabels.add(bucket.label.trim().toLowerCase()),
      (!Number.isSafeInteger(bucket.from) ||
        bucket.from !== expectedBucketStart ||
        (bucketIndex === agingConfiguration.buckets.length - 1
          ? bucket.to !== null
          : !Number.isSafeInteger(bucket.to) ||
            bucket.to < bucket.from ||
            bucket.to > 3652424)) &&
        throwErrorMessage(
          "Buckets must cover every nonnegative day exactly once: start at 0, use contiguous inclusive integer bounds, and leave only the final upper bound open.",
        ),
      (expectedBucketStart = bucket.to === null ? 0 : bucket.to + 1));
  const materialityReportText = parseMaterialityReportText(
    workbook,
    JSON.stringify({
      schemaVersion: 1,
      kind: "materiality-screen",
      request: defaultMaterialityConfig(agingConfiguration),
    }),
    scopeOptions,
  );
  return {
    sheetId: materialityReportText.sheetId,
    range: materialityReportText.range,
    headers: materialityReportText.headers,
    amountColumn: materialityReportText.amountColumn,
    currencyColumn: materialityReportText.currencyColumn,
    dateColumn: materialityReportText.dateColumn,
    currency: materialityReportText.currency,
    dateBasis: agingConfiguration.dateBasis,
    dateSystem: agingConfiguration.dateSystem,
    asOf: agingConfiguration.asOf,
    buckets: agingConfiguration.buckets.map((bucket) => ({
      ...bucket,
      label: bucket.label.trim(),
    })),
  };
}
function addDecimalAmounts(amounts) {
  const decimalParts = amounts.map((value) => {
      const [mantissa, exponentText = "0"] = String(value)
          .toLowerCase()
          .split("e"),
        [integerDigits, fractionDigits = ""] = mantissa.split(".");
      return {
        n: BigInt(integerDigits + fractionDigits),
        scale: fractionDigits.length - Number(exponentText),
      };
    }),
    maxScale = Math.max(
      0,
      ...decimalParts.map((decimalPart) => decimalPart.scale),
    );
  let scaledTotal = decimalParts.reduce(
    (sum, decimalEntry) =>
      sum + decimalEntry.n * 10n ** BigInt(maxScale - decimalEntry.scale),
    0n,
  );
  if (scaledTotal === 0n) return "0";
  const sign = scaledTotal < 0n ? "-" : "";
  scaledTotal < 0n && (scaledTotal = -scaledTotal);
  const digits = scaledTotal.toString().padStart(maxScale + 1, "0");
  return (
    sign +
    (maxScale
      ? `${digits.slice(0, -maxScale)}.${digits.slice(-maxScale)}`.replace(
          /\.?0+$/,
          "",
        )
      : digits)
  );
}
function assertAgingControlTotal(exactAmount) {
  const numericAmount = Number(exactAmount);
  return Number.isFinite(numericAmount)
    ? numericAmount
    : throwErrorMessage(
        "The aging control total exceeds finite numeric capacity. Reduce the selected population or magnitude.",
      );
}
function buildAgingSchedule(workbook, agingRequest, options = {}) {
  const agingConfiguration = normalizeAgingConfiguration(
      workbook,
      agingRequest,
      options,
    ),
    materialityRequest = normalizeMaterialityRequest(
      workbook,
      defaultMaterialityConfig(agingConfiguration),
      options,
    ),
    firstCol = parseRangeAddress(agingConfiguration.range).firstCol,
    agingIsoDate = parseAgingIsoDate(agingConfiguration.asOf),
    rows = materialityRequest.rows.map((row) => {
      const rawAmount = row.values[agingConfiguration.amountColumn - firstCol],
        rawCurrency = row.values[agingConfiguration.currencyColumn - firstCol],
        amount =
          typeof rawAmount == "number" && Number.isFinite(rawAmount)
            ? rawAmount
            : null,
        currency =
          typeof rawCurrency == "string" && /^[A-Za-z]{3}$/.test(rawCurrency)
            ? rawCurrency.toUpperCase()
            : null,
        rawDate = row.values[agingConfiguration.dateColumn - firstCol],
        agingDateInput = normalizeAgingDateInput(
          rawDate,
          agingConfiguration.dateSystem,
        ),
        ageDays =
          agingDateInput === null
            ? null
            : agingIsoDate - parseAgingIsoDate(agingDateInput),
        exceptions = [],
        reasons = [];
      (rawDate === null || rawDate === ""
        ? (exceptions.push("missing-date"),
          reasons.push("Date is missing; no date is inferred."))
        : agingDateInput === null
          ? (exceptions.push("invalid-date"),
            reasons.push(
              `Date must be strict ISO text or an integer ${agingConfiguration.dateSystem}-system calendar serial; time fractions and fictitious 1900-02-29 are unsupported.`,
            ))
          : ageDays < 0 &&
            (exceptions.push("future-date"),
            reasons.push(
              "Date is after the fixed as-of date; future items are excluded and separately counted.",
            )),
        amount === null &&
          (exceptions.push("invalid-amount"),
          reasons.push(
            "Amount is missing or is not a finite number; numeric text is not coerced.",
          )),
        currency !== agingConfiguration.currency &&
          (exceptions.push("unsupported-currency"),
          reasons.push(
            `Currency is missing, invalid or differs from declared ${agingConfiguration.currency}; no conversion or cross-currency aggregation is performed.`,
          )));
      const bucket = exceptions.length
        ? null
        : agingConfiguration.buckets.findIndex(
            (bucket) =>
              ageDays >= bucket.from &&
              (bucket.to === null || ageDays <= bucket.to),
          );
      return (
        bucket === -1 &&
          throwErrorMessage(
            "The configured aging buckets do not cover a valid source row.",
          ),
        {
          sourceRow: row.sourceRow,
          sourceRange: row.sourceRange,
          values: structuredClone(row.values),
          cells: structuredClone(row.cells),
          amount: amount,
          currency: currency,
          date: agingDateInput,
          ageDays: ageDays,
          bucket: bucket,
          disposition: exceptions.length ? "excluded" : "aged",
          exceptions: exceptions,
          reasons: reasons,
        }
      );
    }),
    knownCurrencyRows = rows.filter(
      (row) =>
        row.amount !== null && row.currency === agingConfiguration.currency,
    ),
    agedRows = rows.filter((row) => row.disposition === "aged"),
    excludedKnownRows = knownCurrencyRows.filter(
      (knownRow) => knownRow.disposition === "excluded",
    ),
    sumExactAmounts = (selectedRows) =>
      addDecimalAmounts(selectedRows.map((entry) => entry.amount)),
    exactKnownSourceAmount = sumExactAmounts(knownCurrencyRows),
    exactAgedAmount = sumExactAmounts(agedRows),
    exactExcludedKnownAmount = sumExactAmounts(excludedKnownRows),
    buckets = agingConfiguration.buckets.map((bucket, bucketIndex) => {
      const bucketRows = agedRows.filter(
          (agedRow) => agedRow.bucket === bucketIndex,
        ),
        exactAmount = sumExactAmounts(bucketRows);
      return {
        ...bucket,
        count: bucketRows.length,
        amount: assertAgingControlTotal(exactAmount),
        exactAmount: exactAmount,
      };
    }),
    countExceptions = (exceptionType) =>
      rows.filter((row) => row.exceptions.includes(exceptionType)).length,
    schedule = {
      schemaVersion: 1,
      kind: "aging-schedule",
      algorithm: "ledgerglass-aging-1",
      workbookId: workbook.id,
      revision: workbook.revision,
      sheetId: materialityRequest.sheetId,
      sheetName: materialityRequest.sheetName,
      generatedAt: new Date().toISOString(),
      request: structuredClone(agingConfiguration),
      context: structuredClone(materialityRequest.context),
      status:
        agedRows.length === rows.length
          ? "complete"
          : "complete-with-exclusions",
      headers: structuredClone(materialityRequest.headers),
      rows: rows,
      buckets: buckets,
      counts: {
        population: rows.length,
        aged: agedRows.length,
        excluded: rows.length - agedRows.length,
        missingDates: countExceptions("missing-date"),
        invalidDates: countExceptions("invalid-date"),
        futureDates: countExceptions("future-date"),
        unsupportedCurrencies: countExceptions("unsupported-currency"),
        missingCurrencies: rows.filter((row) => {
          const currencyValue =
            row.values[agingConfiguration.currencyColumn - firstCol];
          return currencyValue === null || currencyValue === "";
        }).length,
        invalidAmounts: countExceptions("invalid-amount"),
      },
      totals: {
        currency: agingConfiguration.currency,
        knownSourceAmount: assertAgingControlTotal(exactKnownSourceAmount),
        agedAmount: assertAgingControlTotal(exactAgedAmount),
        excludedKnownAmount: assertAgingControlTotal(exactExcludedKnownAmount),
        unknownOrOtherCurrencyRows: rows.length - knownCurrencyRows.length,
        exactKnownSourceAmount: exactKnownSourceAmount,
        exactAgedAmount: exactAgedAmount,
        exactExcludedKnownAmount: exactExcludedKnownAmount,
        reconciled: true,
      },
      notes: [
        "Inclusive bucket endpoints cover every nonnegative day exactly once; the final bucket is open-ended. Age is the UTC calendar-day difference from the chosen due/document date to the fixed as-of date.",
        "Every source data row is either aged once or excluded once. Exception categories may overlap. Future, missing and invalid dates are distinct categories. Negative and zero balances retain their sign.",
        "Only the explicitly declared currency is summed. Missing, invalid and other currencies are separately counted exceptions; no foreign-exchange conversion is performed.",
        "Exact decimal controls sum each stored/computed number's shortest decimal representation. Bucket amounts equal the aged population; aged plus known-currency exclusions equals the known-currency source. Unknown amounts and other currencies are outside these totals.",
        "Application dates and imported calendar dates use the 1900 system. Select 1904 only for raw, unconverted 1904 serials; choosing a different serial system never changes ISO dates. ISO years 0001–9999 are accepted; year 0000 is an invalid calendar date. Fractional serials are exceptions, not silently truncated.",
        "Selected amount/currency/date formulas use the existing authorized static scalar dependency closure (1,000 cells, depth 8, 5,000 scanned formulas); volatile, array, dynamic, hidden or unauthorized dependencies are rejected. Other evidence formulas remain stored text.",
        "Headers alone are omitted automatically. Total/subtotal lines are ordinary data unless excluded from the selected population. Source cells and history remain unchanged. This local snapshot is not signed or immutable evidence.",
      ],
    };
  return (
    new TextEncoder().encode(JSON.stringify(schedule, null, 2)).length >
      agingScheduleLimits.reportBytes &&
      throwErrorMessage(
        "The aging report exceeds 2,000,000 bytes. Select a smaller population.",
      ),
    _z.set(schedule, {
      snapshot: JSON.stringify(workbook),
      context: JSON.stringify(materialityRequest.context),
      request: JSON.stringify(agingConfiguration),
      content: JSON.stringify(schedule),
    }),
    schedule
  );
}
function checkAgingReportFreshness(
  workbook,
  report,
  agingRequest,
  options = {},
) {
  const snapshotRecord = _z.get(report);
  if (!snapshotRecord || snapshotRecord.content !== JSON.stringify(report))
    return {
      current: false,
      disclose: false,
      reason:
        "This aging report is untrusted or changed. Run the schedule again.",
    };
  try {
    const agingConfiguration = normalizeAgingConfiguration(
        workbook,
        agingRequest,
        options,
      ),
      authorizedScope = {
        includeHidden: !!options.includeHidden,
        allowedSheetIds: [
          ...new Set(
            options.allowedSheetIds ?? workbook.sheets.map((sheet) => sheet.id),
          ),
        ].sort(),
      };
    return snapshotRecord.snapshot !== JSON.stringify(workbook) ||
      snapshotRecord.context !== JSON.stringify(authorizedScope)
      ? {
          current: false,
          disclose: false,
          reason:
            "The source or authorized scope changed. Run the aging schedule again.",
        }
      : JSON.stringify(agingConfiguration) !== snapshotRecord.request
        ? {
            current: false,
            disclose: true,
            reason:
              "The aging settings changed. This is the previous result; run the schedule again.",
          }
        : {
            current: true,
            disclose: true,
            reason: "Current aging result for this source and configuration.",
          };
  } catch {
    return {
      current: false,
      disclose: false,
      reason:
        "The source, settings or authorized scope is no longer valid. Run the aging schedule again.",
    };
  }
}
function exportAgingReportJson(workbook, report, agingRequest, options = {}) {
  const freshness = checkAgingReportFreshness(
    workbook,
    report,
    agingRequest,
    options,
  );
  return (
    freshness.current || throwErrorMessage(freshness.reason),
    JSON.stringify(report, null, 2)
  );
}
function loadAgingConfiguration(workbook, report, options = {}) {
  return (
    (!report ||
      typeof report != "object" ||
      Array.isArray(report) ||
      report.schemaVersion !== 1 ||
      report.kind !== "aging-schedule") &&
      throwErrorMessage("Choose a supported aging configuration or report."),
    normalizeAgingConfiguration(workbook, report.request, options)
  );
}
const listRangeColumnIndexes = (rangeText) => {
    try {
      const rangeAddress = parseRangeAddress(rangeText);
      return rangeAddress.lastCol - rangeAddress.firstCol < 500
        ? Array.from(
            { length: rangeAddress.lastCol - rangeAddress.firstCol + 1 },
            (t, columnOffset) => rangeAddress.firstCol + columnOffset,
          )
        : [];
    } catch {
      return [];
    }
  },
  defaultAgingBuckets = [
    { label: "0–30 days", from: 0, to: 30 },
    { label: "31–60 days", from: 31, to: 60 },
    { label: "61–90 days", from: 61, to: 90 },
    { label: "91+ days", from: 91, to: null },
  ];
function AgingScheduleScreen(props) {
  const { text, language } = useWorkbookMessages(),
    cancellableOperation = useCancellableOperation(),
    titleId = ne.useId(),
    dialogRef = ne.useRef(null),
    isMountedRef = ne.useRef(true),
    requestCounterRef = ne.useRef(0),
    requestKeyRef = ne.useRef(""),
    rangeColumnIndexes = listRangeColumnIndexes(props.range),
    [range, setRange] = ne.useState(props.range),
    [hasHeaders, setHasHeaders] = ne.useState(true),
    [amountColumn, setAmountColumn] = ne.useState(rangeColumnIndexes[0] ?? 0),
    [currencyColumn, setCurrencyColumn] = ne.useState(
      rangeColumnIndexes[1] ?? 1,
    ),
    [dateColumn, setDateColumn] = ne.useState(rangeColumnIndexes[2] ?? 2),
    [currency, setCurrency] = ne.useState("USD"),
    [dateBasis, setDateBasis] = ne.useState("due"),
    [dateSystem, setDateSystem] = ne.useState(
      props.workbook.datePolicy?.dateSystem ?? "1900",
    ),
    [asOf, setAsOf] = ne.useState(new Date().toISOString().slice(0, 10)),
    [buckets, setBuckets] = ne.useState(structuredClone(defaultAgingBuckets)),
    [checked, setChecked] = ne.useState(false),
    [agingSchedule, setAgingSchedule] = ne.useState(null),
    [runSourceKey, setRunSourceKey] = ne.useState(""),
    [errorMessage, setErrorMessage] = ne.useState(""),
    [errorScopeKey, setErrorScopeKey] = ne.useState(""),
    [statusMessage, setStatusMessage] = ne.useState(""),
    [X$1, setPageIndex] = ne.useState(0),
    [reviewFilter, setReviewFilter] = ne.useState("all"),
    agingRequest = {
      sheetId: props.sheetId,
      range: range,
      headers: hasHeaders,
      amountColumn: amountColumn,
      currencyColumn: currencyColumn,
      dateColumn: dateColumn,
      currency: currency,
      dateBasis: dateBasis,
      dateSystem: dateSystem,
      asOf: asOf,
      buckets: buckets,
    },
    agingOptions = {
      readOnly: props.readOnly,
      includeHidden: !!props.allowHidden && checked,
      allowedSheetIds: props.allowedSheetIds,
    },
    sheet = props.workbook.sheets.find((sheet) => sheet.id === props.sheetId),
    sortedSheetIds = [
      ...new Set(
        props.allowedSheetIds ?? props.workbook.sheets.map((sheet) => sheet.id),
      ),
    ].sort(),
    isSourceInContext =
      !!sheet &&
      sortedSheetIds.includes(sheet.id) &&
      (agingOptions.includeHidden || (!sheet.hidden && !sheet.veryHidden)),
    sourceKey = JSON.stringify([props.sheetId, props.range]),
    workbookKey = JSON.stringify([
      props.workbook,
      sortedSheetIds,
      agingOptions.includeHidden,
      sourceKey,
    ]),
    requestKey = JSON.stringify([workbookKey, agingRequest]);
  requestKeyRef.current = requestKey;
  const freshness = agingSchedule
      ? runSourceKey !== sourceKey
        ? {
            current: false,
            disclose: false,
            reason:
              "The source selection changed. Run the aging schedule again.",
          }
        : checkAgingReportFreshness(
            props.workbook,
            agingSchedule,
            agingRequest,
            agingOptions,
          )
      : null,
    visibleSchedule = freshness?.disclose ? agingSchedule : null,
    filteredRows =
      visibleSchedule?.rows.filter(
        (row) =>
          reviewFilter === "all" ||
          reviewFilter === row.disposition ||
          row.exceptions.includes(reviewFilter),
      ) ?? [],
    pageCount = Math.max(
      1,
      Math.ceil(filteredRows.length / agingScheduleLimits.pageSize),
    ),
    currentPage = Math.min(X$1, pageCount - 1),
    selectableColumnIndexes = listRangeColumnIndexes(range),
    showError = (caughtError) => {
      (setErrorMessage(
        caughtError instanceof Error
          ? caughtError.message
          : "The aging schedule could not complete. Review the settings and source.",
      ),
        setErrorScopeKey(workbookKey),
        setStatusMessage(""));
    };
  ne.useEffect(() => {
    isMountedRef.current = true;
    const activeElement = document.activeElement;
    return (
      dialogRef.current
        ?.querySelector('[data-ui-label="Aging source range"]')
        ?.focus(),
      () => {
        ((isMountedRef.current = false),
          requestCounterRef.current++,
          activeElement?.isConnected && activeElement.focus());
      }
    );
  }, []);
  const onClick = () => {
      requestCounterRef.current++;
      try {
        (setAgingSchedule(
          buildAgingSchedule(props.workbook, agingRequest, agingOptions),
        ),
          setRunSourceKey(sourceKey),
          setErrorMessage(""),
          setStatusMessage(""),
          setReviewFilter("all"),
          setPageIndex(0));
      } catch (error) {
        (setAgingSchedule(null), showError(error));
      }
    },
    downloadReport = () => {
      try {
        if (!agingSchedule || !freshness?.current)
          throw Error("Run a current aging schedule before downloading.");
        const reportJson = exportAgingReportJson(
            props.workbook,
            agingSchedule,
            agingRequest,
            agingOptions,
          ),
          downloadUrl = URL.createObjectURL(
            new Blob([reportJson], { type: "application/json" }),
          ),
          downloadLink = document.createElement("a");
        ((downloadLink.href = downloadUrl),
          (downloadLink.download = "aging-schedule.json"),
          downloadLink.click(),
          URL.revokeObjectURL(downloadUrl),
          setStatusMessage(
            "Aging report downloaded with settings, all source rows, exceptions and exact control totals.",
          ),
          setErrorMessage(""));
      } catch (error) {
        showError(error);
      }
    },
    loadConfigurationFile = async (configFile) => {
      if (!configFile) return;
      const startRequestKey = requestKey,
        requestId = ++requestCounterRef.current,
        isCurrentRequest = () =>
          isMountedRef.current &&
          requestId === requestCounterRef.current &&
          requestKeyRef.current === startRequestKey;
      try {
        if (configFile.size > agingScheduleLimits.reportBytes)
          throw Error(
            "Choose an aging configuration/report no larger than 2,000,000 bytes.",
          );
        const importedConfig = await importNativeJsonFile(
          configFile,
          cancellableOperation.begin(),
        );
        if (!isCurrentRequest()) return;
        const agingConfiguration = loadAgingConfiguration(
          props.workbook,
          importedConfig,
          agingOptions,
        );
        if (agingConfiguration.sheetId !== props.sheetId)
          throw Error(
            "Open the saved configuration's source worksheet before loading it.",
          );
        (setRange(agingConfiguration.range),
          setHasHeaders(agingConfiguration.headers),
          setAmountColumn(agingConfiguration.amountColumn),
          setCurrencyColumn(agingConfiguration.currencyColumn),
          setDateColumn(agingConfiguration.dateColumn),
          setCurrency(agingConfiguration.currency),
          setDateBasis(agingConfiguration.dateBasis),
          setDateSystem(agingConfiguration.dateSystem),
          setAsOf(agingConfiguration.asOf),
          setBuckets(agingConfiguration.buckets),
          setAgingSchedule(null),
          setErrorMessage(""),
          setStatusMessage(
            "Aging configuration loaded. Run it against the current source; saved conclusions are not reused.",
          ));
      } catch (error) {
        isCurrentRequest() && showError(toFileProcessingError(error));
      }
    },
    updateBucket = (indexToUpdate, changes) =>
      setBuckets((bucketsToUpdate) =>
        bucketsToUpdate.map((bucketEntry, entryIndex) =>
          entryIndex === indexToUpdate
            ? { ...bucketEntry, ...changes }
            : bucketEntry,
        ),
      );
  return i.jsx("div", {
    className: "aging-backdrop",
    onMouseDown: (event) => {
      event.target === event.currentTarget && props.onClose();
    },
    children: i.jsxs("section", {
      className: "aging-dialog",
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": `${titleId}-title`,
      ref: dialogRef,
      onKeyDown: (event) => {
        var lastFocusable, firstFocusable;
        if (
          (event.stopPropagation(),
          event.key === "Escape" && (event.preventDefault(), props.onClose()),
          event.key === "Tab")
        ) {
          const focusableElements = Array.from(
            dialogRef.current?.querySelectorAll(
              'button:not(:disabled), input:not(:disabled), select:not(:disabled), summary, [tabindex="0"]',
            ) ?? [],
          ).filter(
            (element) =>
              !element.closest("details:not([open])") ||
              element.tagName === "SUMMARY",
          );
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
                  id: `${titleId}-title`,
                  children: i.jsx(LocalizedText, {
                    children: "Aging schedule",
                  }),
                }),
                i.jsx("p", {
                  children: i.jsx(LocalizedText, {
                    children:
                      "Age receivables or payables from a due or document date using your fixed as-of date and inclusive buckets.",
                  }),
                }),
              ],
            }),
            i.jsx("button", {
              "aria-label": text("Close aging schedule"),
              onClick: props.onClose,
              lang: language("Close aging schedule"),
              "data-ui-label": "Close aging schedule",
              children: i.jsx(X, { size: 18 }),
            }),
          ],
        }),
        i.jsx("p", {
          children: i.jsx(LocalizedText, {
            children:
              "This schedule reads the source without editing cells or history. Protected and read-only workbooks support inspection.",
          }),
        }),
        i.jsxs("div", {
          className: "aging-fields",
          children: [
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Source worksheet" }),
                i.jsx("strong", {
                  children: isSourceInContext
                    ? sheet.name
                    : "Outside current context",
                }),
              ],
            }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Source range" }),
                i.jsx("input", {
                  "aria-label": text("Aging source range"),
                  value: isSourceInContext ? range : "",
                  onChange: (event) => setRange(event.target.value),
                  lang: language("Aging source range"),
                  "data-ui-label": "Aging source range",
                }),
              ],
            }),
            i.jsxs("label", {
              className: "aging-check",
              children: [
                i.jsx("input", {
                  type: "checkbox",
                  checked: hasHeaders,
                  onChange: (event) => setHasHeaders(event.target.checked),
                }),
                i.jsx(LocalizedText, {
                  children: "First row contains headers",
                }),
              ],
            }),
            [
              ["Amount column", amountColumn, setAmountColumn],
              ["Currency column", currencyColumn, setCurrencyColumn],
              ["Date column", dateColumn, setDateColumn],
            ].map(([fieldLabel, selectedColumn, setSelectedColumn]) =>
              i.jsxs(
                "label",
                {
                  children: [
                    fieldLabel,
                    i.jsx("select", {
                      "aria-label": fieldLabel,
                      value: selectedColumn,
                      onChange: (event) =>
                        setSelectedColumn(Number(event.target.value)),
                      children: selectableColumnIndexes.map((columnIndex) =>
                        i.jsx(
                          "option",
                          {
                            value: columnIndex,
                            children: columnIndexToLetters(columnIndex),
                          },
                          columnIndex,
                        ),
                      ),
                    }),
                  ],
                },
                fieldLabel,
              ),
            ),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Declared currency" }),
                i.jsx("input", {
                  "aria-label": text("Declared currency"),
                  maxLength: 3,
                  value: currency,
                  onChange: (event) => setCurrency(event.target.value),
                  lang: language("Declared currency"),
                  "data-ui-label": "Declared currency",
                }),
              ],
            }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Date basis" }),
                i.jsxs("select", {
                  "aria-label": text("Date basis"),
                  value: dateBasis,
                  onChange: (event) => setDateBasis(event.target.value),
                  lang: language("Date basis"),
                  "data-ui-label": "Date basis",
                  children: [
                    i.jsx("option", { value: "due", children: "Due date" }),
                    i.jsx("option", {
                      value: "document",
                      children: "Document date",
                    }),
                  ],
                }),
              ],
            }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Fixed as-of date" }),
                i.jsx("input", {
                  "aria-label": text("Fixed as-of date"),
                  type: "date",
                  value: asOf,
                  onChange: (event) => setAsOf(event.target.value),
                  lang: language("Fixed as-of date"),
                  "data-ui-label": "Fixed as-of date",
                }),
              ],
            }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, {
                  children: "Numeric serial date system",
                }),
                i.jsxs("select", {
                  "aria-label": text("Numeric serial date system"),
                  value: dateSystem,
                  onChange: (event) => setDateSystem(event.target.value),
                  lang: language("Numeric serial date system"),
                  "data-ui-label": "Numeric serial date system",
                  children: [
                    i.jsx("option", {
                      value: "1900",
                      children: "1900 — application and imported dates",
                    }),
                    i.jsx("option", {
                      value: "1904",
                      children: "1904 — raw unconverted serials",
                    }),
                  ],
                }),
              ],
            }),
          ],
        }),
        i.jsxs("fieldset", {
          children: [
            i.jsx("legend", {
              children: i.jsx(LocalizedText, {
                children: "Inclusive aging buckets",
              }),
            }),
            i.jsx("p", {
              children: i.jsx(LocalizedText, {
                children:
                  "Start at day 0. Bounds must be contiguous without overlap; the last bucket is open-ended. Future dates are separate exceptions.",
              }),
            }),
            i.jsx("div", {
              className: "aging-scroll",
              children: i.jsxs("table", {
                children: [
                  i.jsx("thead", {
                    children: i.jsxs("tr", {
                      children: [
                        i.jsx("th", {
                          children: i.jsx(LocalizedText, { children: "Label" }),
                        }),
                        i.jsx("th", {
                          children: i.jsx(LocalizedText, {
                            children: "From day (inclusive)",
                          }),
                        }),
                        i.jsx("th", {
                          children: i.jsx(LocalizedText, {
                            children: "To day (inclusive)",
                          }),
                        }),
                        i.jsx("th", {
                          children: i.jsx(LocalizedText, {
                            children: "Action",
                          }),
                        }),
                      ],
                    }),
                  }),
                  i.jsx("tbody", {
                    children: buckets.map((bucket, bucketIndex) =>
                      i.jsxs(
                        "tr",
                        {
                          children: [
                            i.jsx("td", {
                              children: i.jsx("input", {
                                "aria-label": `Bucket ${bucketIndex + 1} label`,
                                value: bucket.label,
                                maxLength: 80,
                                onChange: (event) =>
                                  updateBucket(bucketIndex, {
                                    label: event.target.value,
                                  }),
                              }),
                            }),
                            i.jsx("td", {
                              children: i.jsx(LocalizedNumberField, {
                                "aria-label": `Bucket ${bucketIndex + 1} from`,
                                min: "0",
                                step: "1",
                                value: Number.isFinite(bucket.from)
                                  ? bucket.from
                                  : "",
                                onValueChange: (fromInput) =>
                                  updateBucket(bucketIndex, {
                                    from:
                                      fromInput === ""
                                        ? NaN
                                        : Number(fromInput),
                                  }),
                              }),
                            }),
                            i.jsx("td", {
                              children: i.jsx(LocalizedNumberField, {
                                allowEmpty: true,
                                "aria-label": `Bucket ${bucketIndex + 1} to`,
                                min: "0",
                                step: "1",
                                placeholder: text("Open-ended"),
                                value:
                                  bucket.to === null
                                    ? ""
                                    : Number.isFinite(bucket.to)
                                      ? bucket.to
                                      : "",
                                onValueChange: (toInput) =>
                                  updateBucket(bucketIndex, {
                                    to: toInput === "" ? null : Number(toInput),
                                  }),
                                lang: language("Open-ended"),
                              }),
                            }),
                            i.jsx("td", {
                              children: i.jsx("button", {
                                "aria-label": `Remove bucket ${bucketIndex + 1}`,
                                disabled: buckets.length === 1,
                                onClick: () =>
                                  setBuckets((currentBuckets) =>
                                    currentBuckets.filter(
                                      ($e, filterIndex) =>
                                        filterIndex !== bucketIndex,
                                    ),
                                  ),
                                children: i.jsx(LocalizedText, {
                                  children: "Remove",
                                }),
                              }),
                            }),
                          ],
                        },
                        bucketIndex,
                      ),
                    ),
                  }),
                ],
              }),
            }),
            i.jsx("button", {
              disabled: buckets.length >= agingScheduleLimits.buckets,
              onClick: () =>
                setBuckets((previousBuckets) => {
                  const lastBucket = previousBuckets.at(-1),
                    lastBucketEnd = lastBucket.to ?? lastBucket.from + 29;
                  return [
                    ...previousBuckets.slice(0, -1),
                    { ...lastBucket, to: lastBucketEnd },
                    {
                      label: `${lastBucketEnd + 1}+ days`,
                      from: lastBucketEnd + 1,
                      to: null,
                    },
                  ];
                }),
              children: i.jsx(LocalizedText, { children: "Add bucket" }),
            }),
            i.jsx("button", {
              onClick: () => setBuckets(structuredClone(defaultAgingBuckets)),
              children: i.jsx(LocalizedText, { children: "Reset buckets" }),
            }),
          ],
        }),
        props.allowHidden &&
          i.jsxs("label", {
            className: "aging-check",
            children: [
              i.jsx("input", {
                "aria-label": text("Include hidden content"),
                type: "checkbox",
                checked: checked,
                onChange: (event) => setChecked(event.target.checked),
                lang: language("Include hidden content"),
                "data-ui-label": "Include hidden content",
              }),
              i.jsx(LocalizedText, {
                children:
                  "Explicitly include hidden source and dependency cells within permitted host context",
              }),
            ],
          }),
        i.jsx("p", {
          children: i.jsx(LocalizedText, {
            children:
              "Up to 500 selected cells. Dates use strict ISO calendar text or whole-day serials in the selected system. Invalid, missing, future and other-currency rows are counted separately. The declared currency alone is totaled; no conversion is performed.",
          }),
        }),
        i.jsxs("div", {
          className: "aging-actions",
          children: [
            i.jsx("button", {
              className: "aging-primary",
              disabled: !isSourceInContext,
              onClick: onClick,
              children: i.jsx(LocalizedText, {
                children: "Run aging schedule",
              }),
            }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, {
                  children: "Load saved configuration/report",
                }),
                i.jsx("input", {
                  "aria-label": text("Load aging configuration"),
                  type: "file",
                  accept: "application/json,.json",
                  onChange: (event) => {
                    const selectedFile = event.target.files?.[0];
                    ((event.target.value = ""),
                      loadConfigurationFile(selectedFile));
                  },
                  lang: language("Load aging configuration"),
                  "data-ui-label": "Load aging configuration",
                }),
              ],
            }),
            i.jsx("button", {
              disabled: !freshness?.current,
              onClick: downloadReport,
              children: i.jsx(LocalizedText, {
                children: "Download current aging report",
              }),
            }),
          ],
        }),
        errorMessage &&
          errorScopeKey === workbookKey &&
          i.jsx("p", {
            role: "alert",
            className: "aging-error",
            children: errorMessage,
          }),
        statusMessage &&
          i.jsx("p", { role: "status", children: statusMessage }),
        agingSchedule &&
          freshness &&
          i.jsx("p", {
            role: "status",
            className: freshness.current ? "" : "aging-stale",
            children: freshness.current
              ? freshness.reason
              : `Stale result. ${freshness.reason}`,
          }),
        visibleSchedule &&
          i.jsxs("section", {
            className: "aging-result",
            "aria-label": text("Aging schedule results"),
            lang: language("Aging schedule results"),
            "data-ui-label": "Aging schedule results",
            children: [
              i.jsx("h3", {
                children:
                  visibleSchedule.status === "complete"
                    ? "Schedule complete"
                    : "Complete with exclusions — review every excluded row",
              }),
              i.jsxs("p", {
                children: [
                  visibleSchedule.request.dateBasis === "due"
                    ? "Due date"
                    : "Document date",
                  " ",
                  i.jsx(LocalizedText, { children: "basis · as of " }),
                  visibleSchedule.request.asOf,
                  " · ",
                  visibleSchedule.request.dateSystem,
                  " ",
                  i.jsx(LocalizedText, { children: "serial system ·" }),
                  " ",
                  visibleSchedule.request.currency,
                ],
              }),
              i.jsxs("p", {
                children: [
                  visibleSchedule.counts.population,
                  " ",
                  i.jsx(LocalizedText, { children: "source rows ·" }),
                  " ",
                  visibleSchedule.counts.aged,
                  " ",
                  i.jsx(LocalizedText, { children: "aged · " }),
                  visibleSchedule.counts.excluded,
                  " ",
                  i.jsx(LocalizedText, { children: "excluded" }),
                ],
              }),
              i.jsxs("p", {
                children: [
                  visibleSchedule.counts.missingDates,
                  " ",
                  i.jsx(LocalizedText, { children: "missing dates ·" }),
                  " ",
                  visibleSchedule.counts.invalidDates,
                  " ",
                  i.jsx(LocalizedText, { children: "invalid dates ·" }),
                  " ",
                  visibleSchedule.counts.futureDates,
                  " ",
                  i.jsx(LocalizedText, { children: "future dates ·" }),
                  " ",
                  visibleSchedule.counts.unsupportedCurrencies,
                  " ",
                  i.jsx(LocalizedText, {
                    children: "unsupported currencies (including ",
                  }),
                  visibleSchedule.counts.missingCurrencies,
                  " ",
                  i.jsx(LocalizedText, { children: "missing currencies) · " }),
                  visibleSchedule.counts.invalidAmounts,
                  " ",
                  i.jsx(LocalizedText, {
                    children:
                      "invalid amounts. Exception categories may overlap.",
                  }),
                ],
              }),
              i.jsxs("table", {
                "aria-label": text("Aging bucket totals"),
                lang: language("Aging bucket totals"),
                "data-ui-label": "Aging bucket totals",
                children: [
                  i.jsx("thead", {
                    children: i.jsxs("tr", {
                      children: [
                        i.jsx("th", {
                          children: i.jsx(LocalizedText, {
                            children: "Bucket",
                          }),
                        }),
                        i.jsx("th", {
                          children: i.jsx(LocalizedText, {
                            children: "Inclusive days",
                          }),
                        }),
                        i.jsx("th", {
                          children: i.jsx(LocalizedText, { children: "Rows" }),
                        }),
                        i.jsxs("th", {
                          children: [
                            i.jsx(LocalizedText, { children: "Signed " }),
                            visibleSchedule.totals.currency,
                            " ",
                            i.jsx(LocalizedText, { children: "total" }),
                          ],
                        }),
                      ],
                    }),
                  }),
                  i.jsx("tbody", {
                    children: visibleSchedule.buckets.map(
                      (bucket, resultBucketIndex) =>
                        i.jsxs(
                          "tr",
                          {
                            children: [
                              i.jsx("th", {
                                scope: "row",
                                children: bucket.label,
                              }),
                              i.jsxs("td", {
                                children: [
                                  bucket.from,
                                  "–",
                                  bucket.to ?? "unbounded",
                                ],
                              }),
                              i.jsx("td", { children: bucket.count }),
                              i.jsx("td", { children: bucket.exactAmount }),
                            ],
                          },
                          resultBucketIndex,
                        ),
                    ),
                  }),
                ],
              }),
              i.jsxs("p", {
                children: [
                  i.jsx(LocalizedText, { children: "Exact signed " }),
                  visibleSchedule.totals.currency,
                  " ",
                  i.jsx(LocalizedText, {
                    children: "control totals: known source ",
                  }),
                  visibleSchedule.totals.exactKnownSourceAmount,
                  " ",
                  i.jsx(LocalizedText, { children: "= aged" }),
                  " ",
                  visibleSchedule.totals.exactAgedAmount,
                  " ",
                  i.jsx(LocalizedText, { children: "+ excluded known amount" }),
                  " ",
                  visibleSchedule.totals.exactExcludedKnownAmount,
                  ".",
                  " ",
                  visibleSchedule.totals.unknownOrOtherCurrencyRows,
                  " ",
                  i.jsx(LocalizedText, {
                    children:
                      "rows have unknown amounts or another/unknown currency and are outside these totals.",
                  }),
                ],
              }),
              i.jsxs("p", {
                children: [
                  i.jsx(LocalizedText, { children: "Source: " }),
                  visibleSchedule.sheetName,
                  "!",
                  visibleSchedule.request.range,
                  " ",
                  i.jsx(LocalizedText, { children: "· workbook " }),
                  visibleSchedule.workbookId,
                  " ",
                  i.jsx(LocalizedText, { children: "· revision" }),
                  " ",
                  visibleSchedule.revision,
                  " · ",
                  visibleSchedule.algorithm,
                ],
              }),
              i.jsxs("label", {
                children: [
                  i.jsx(LocalizedText, { children: "Rows to review" }),
                  i.jsxs("select", {
                    "aria-label": text("Aging rows to review"),
                    value: reviewFilter,
                    onChange: (event) => {
                      (setReviewFilter(event.target.value), setPageIndex(0));
                    },
                    lang: language("Aging rows to review"),
                    "data-ui-label": "Aging rows to review",
                    children: [
                      i.jsx("option", {
                        value: "all",
                        children: "All source rows",
                      }),
                      i.jsx("option", { value: "aged", children: "Aged" }),
                      i.jsx("option", {
                        value: "excluded",
                        children: "Excluded",
                      }),
                      i.jsx("option", {
                        value: "missing-date",
                        children: "Missing date",
                      }),
                      i.jsx("option", {
                        value: "invalid-date",
                        children: "Invalid date",
                      }),
                      i.jsx("option", {
                        value: "future-date",
                        children: "Future date",
                      }),
                      i.jsx("option", {
                        value: "unsupported-currency",
                        children: "Unsupported currency",
                      }),
                      i.jsx("option", {
                        value: "invalid-amount",
                        children: "Invalid amount",
                      }),
                    ],
                  }),
                ],
              }),
              i.jsx("div", {
                className: "aging-scroll",
                children: i.jsxs("table", {
                  "aria-label": text("Aging source rows"),
                  lang: language("Aging source rows"),
                  "data-ui-label": "Aging source rows",
                  children: [
                    i.jsx("thead", {
                      children: i.jsxs("tr", {
                        children: [
                          i.jsx("th", {
                            children: i.jsx(LocalizedText, {
                              children: "Source row / range",
                            }),
                          }),
                          i.jsx("th", {
                            children: i.jsx(LocalizedText, {
                              children: "Date / age",
                            }),
                          }),
                          i.jsx("th", {
                            children: i.jsx(LocalizedText, {
                              children: "Amount / currency",
                            }),
                          }),
                          i.jsx("th", {
                            children: i.jsx(LocalizedText, {
                              children: "Bucket / disposition",
                            }),
                          }),
                          i.jsx("th", {
                            children: i.jsx(LocalizedText, {
                              children: "Exceptions",
                            }),
                          }),
                          visibleSchedule.headers.map((header, headerIndex) =>
                            i.jsxs(
                              "th",
                              {
                                children: [
                                  header,
                                  " (",
                                  columnIndexToLetters(
                                    parseRangeAddress(
                                      visibleSchedule.request.range,
                                    ).firstCol + headerIndex,
                                  ),
                                  ")",
                                ],
                              },
                              headerIndex,
                            ),
                          ),
                        ],
                      }),
                    }),
                    i.jsx("tbody", {
                      children: filteredRows
                        .slice(currentPage * 20, currentPage * 20 + 20)
                        .map((filteredRow) =>
                          i.jsxs(
                            "tr",
                            {
                              children: [
                                i.jsxs("th", {
                                  scope: "row",
                                  children: [
                                    filteredRow.sourceRow,
                                    " · ",
                                    filteredRow.sourceRange,
                                  ],
                                }),
                                i.jsxs("td", {
                                  children: [
                                    filteredRow.date ?? "Invalid or missing",
                                    " /",
                                    " ",
                                    filteredRow.ageDays ?? "Unavailable",
                                  ],
                                }),
                                i.jsxs("td", {
                                  children: [
                                    filteredRow.amount ?? "Invalid",
                                    " /",
                                    " ",
                                    filteredRow.currency ?? "Invalid",
                                  ],
                                }),
                                i.jsx("td", {
                                  children:
                                    filteredRow.bucket === null
                                      ? "Excluded"
                                      : visibleSchedule.buckets[
                                          filteredRow.bucket
                                        ].label,
                                }),
                                i.jsx("td", {
                                  children:
                                    filteredRow.reasons.join(" ") ||
                                    "Aged once in the configured bucket.",
                                }),
                                filteredRow.cells.map((cell, cellIndex) =>
                                  i.jsx(
                                    "td",
                                    {
                                      children: i.jsxs("details", {
                                        children: [
                                          i.jsx("summary", {
                                            children:
                                              JSON.stringify(
                                                filteredRow.values[cellIndex],
                                              ) ?? "null",
                                          }),
                                          i.jsx("p", {
                                            children: i.jsx(LocalizedText, {
                                              children: "Stored cell evidence",
                                            }),
                                          }),
                                          i.jsx("pre", {
                                            children: JSON.stringify(
                                              cell,
                                              null,
                                              2,
                                            ),
                                          }),
                                        ],
                                      }),
                                    },
                                    cellIndex,
                                  ),
                                ),
                              ],
                            },
                            filteredRow.sourceRow,
                          ),
                        ),
                    }),
                  ],
                }),
              }),
              i.jsxs("nav", {
                "aria-label": text("Aging result pages"),
                lang: language("Aging result pages"),
                "data-ui-label": "Aging result pages",
                children: [
                  i.jsx("button", {
                    disabled: currentPage === 0,
                    onClick: () => setPageIndex(currentPage - 1),
                    children: i.jsx(LocalizedText, {
                      children: "Previous rows",
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
                      filteredRows.length,
                      " ",
                      i.jsx(LocalizedText, { children: "rows in this view" }),
                    ],
                  }),
                  i.jsx("button", {
                    disabled: currentPage + 1 >= pageCount,
                    onClick: () => setPageIndex(currentPage + 1),
                    children: i.jsx(LocalizedText, { children: "Next rows" }),
                  }),
                ],
              }),
              i.jsxs("details", {
                children: [
                  i.jsx("summary", {
                    children: i.jsx(LocalizedText, {
                      children: "Aging rules and limitations",
                    }),
                  }),
                  visibleSchedule.notes.map((note) =>
                    i.jsx("p", { children: note }, note),
                  ),
                ],
              }),
            ],
          }),
        i.jsx("footer", {
          children: i.jsx("button", {
            onClick: props.onClose,
            children: i.jsx(LocalizedText, {
              children: "Close aging schedule",
            }),
          }),
        }),
      ],
    }),
  });
}

export {
  AgingScheduleScreen,
  materialityScreenLimits,
  loadMaterialityConfiguration,
  checkMaterialityReportFreshness,
  requestAssistantAnswer,
  buildMaterialitySourceEvidence,
  exportMaterialityReportJson,
  runMaterialityScreen,
};
export { listRangeColumnIndexes };
