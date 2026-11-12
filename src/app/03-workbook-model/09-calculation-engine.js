import {
  resolveExtensionFunctionCall,
  isWithinNestingLimit,
  createCalculationBudgetTracker,
  resolveLocalDateTime,
  listInstalledExtensionFunctions,
  readMonotonicClockMs,
  calculationFeatureSupport,
  findVolatileFunctionNames,
  workbookUsesNowOrToday,
  workbookHasVolatileFunctions,
} from "../01-schemas-date-functions-and-validation.js";
import {
  fileImportLimits,
  workbookLimits,
} from "../02-text-and-financial-functions.js";
import {
  string,
  number,
  boolean,
  null as _null,
  array,
  literal,
  union,
  enum as _enum,
  unknown,
  object,
} from "zod";
import {
  FB,
  validationLimits,
  DataValidationError,
  validateDataValidationRule,
  isCellInsideRange,
  r5,
} from "./01-cell-styles-and-themes.js";
import {
  FormulaParser,
  matrixResultFunctionNames,
} from "./02-formula-function-evaluators.js";
import { evaluateReusableFormula, y5 } from "./03-defined-names.js";
import {
  builtinFunctionNames,
  findFormulaErrorArg,
  createFormulaError,
  resolveDefinedNameValue,
  buildExternalLinkAliases,
  isFormulaError,
} from "./04-pivot-tables-and-pictures.js";
import {
  hasCalculatorRecord,
  deepFreeze,
  makeFormulaError,
  createWorkbookFormulaEngine,
  reportCellValueMaps,
  getCachedCalculationSnapshot,
  _O,
  carriedCalculatorRecords,
  ik,
  j5,
  getCalculationInputSignature,
  pendingFormulaValue,
  isCalculationRecordCurrent,
  buildReportCellValueMap,
  calculatorRecords,
} from "./05-formula-engine-and-name-operations.js";
import {
  AU,
  nonEmptyLabelSchema,
  reportOutputFormatVersion,
  createBoundedTextSchema,
  columnIndexToLetters,
  shiftFormulaReferences,
  reportHashSchema,
} from "./06-data-model-and-analysis.js";
import {
  findDataValidationsInRange,
  resolveFilteredRowContext,
  normalizeConditionalRules,
  evaluateConditionalFormatRules,
  _t,
  parseRangeAddress,
} from "./07-conditional-rules-worksheet-features.js";
import {
  rewriteFormulaReferences,
  rewriteStructuredReferences,
} from "./08-table-pivot-query-definitions.js";

function evaluateWorkbookFormulas(
  workbook,
  options = {},
  accessPolicy,
  calculationContext,
) {
  const installedExtensionFunctions = listInstalledExtensionFunctions(workbook),
    supported = new Set([
      ...builtinFunctionNames,
      ...installedExtensionFunctions.map(
        (installedExtensionFunction) => installedExtensionFunction.name,
      ),
    ]),
    budgetTracker =
      calculationContext?.budget ?? createCalculationBudgetTracker(workbook),
    clock =
      calculationContext?.dateContext ??
      resolveLocalDateTime(workbook.datePolicy),
    operationalErrors = new Set(),
    referenceDiagnostics = new Map(),
    namedFormulaDiagnostics = new Map(),
    financialDiagnostics = new Map(),
    issue = (errorCode) => (
      [
        "#LIMIT!",
        "#BLOCKED!",
        "#CYCLE!",
        "#ERROR!",
        "#TIMEOUT!",
        "#ITER!",
      ].includes(errorCode) && operationalErrors.add(errorCode),
      createFormulaError(errorCode)
    ),
    filteredRowContext = resolveFilteredRowContext(workbook, options),
    computedCellValues = new Map(),
    spillCellValues = new Map(),
    spillRanges = new Map(),
    spillDiagnostics = new Map(),
    activeFormulaCells = new Set(),
    parsingCells = new Set(),
    readDependencies = new Map(),
    nestedFormulaEngines = [],
    reusableFormulaEngines = [];
  let reusableFormulaDepth = 0;
  const criterionFormulaEngines = [];
  let criterionDepth = 0;
  const hasLetOrLambdaNames = workbook.definedNames?.some((definedName) =>
      /\b(?:_xlfn\.)?(?:LET|LAMBDA)\s*\(/i.test(definedName.refersTo),
    ),
    definedNamePattern = workbook.definedNames?.length
      ? new RegExp(
          `(?:^|[^A-Za-z0-9_.\\\\])(?:${workbook.definedNames.map((definedName) => definedName.name.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")).join("|")})(?![A-Za-z0-9_.\\\\])`,
          "i",
        )
      : undefined,
    isReusableFormula = (textToScan) =>
      /\b(?:_xlfn\.)?(?:LET|LAMBDA)\s*\(/i.test(textToScan) ||
      !!(hasLetOrLambdaNames && definedNamePattern?.test(textToScan)),
    externalLinkAliases = buildExternalLinkAliases(workbook),
    allSheets = [...workbook.sheets, ...externalLinkAliases.sheets],
    sheetsById = new Map(
      allSheets.map((sheetRecord) => [sheetRecord.id, sheetRecord]),
    ),
    sheetsByName = new Map(
      allSheets.map((namedSheetRecord) => [
        namedSheetRecord.name.toLowerCase(),
        namedSheetRecord,
      ]),
    );
  let cellReadCount = 0,
    spilledCellCount = 0;
  const spillCandidateAddresses = new Map(),
    matrixFunctionPattern = new RegExp(
      `\\b(?:${[...matrixResultFunctionNames, "ISFORMULA", "TEXTBEFORE", "TEXTAFTER", ...installedExtensionFunctions.filter((installedExtensionFunction) => installedExtensionFunction.fn.result.type === "matrix").map((extensionFunction) => extensionFunction.name.replaceAll(".", "\\."))].join("|")})\\s*\\(`,
      "i",
    );
  for (const sheet of workbook.sheets)
    spillCandidateAddresses.set(
      sheet.id,
      Object.entries(sheet.cells)
        .filter(
          ([, cell]) =>
            typeof cell.raw == "string" &&
            cell.raw.startsWith("=") &&
            !cell.literal &&
            (isReusableFormula(cell.raw) ||
              matrixFunctionPattern.test(cell.raw) ||
              (/\b(?:IF|IFERROR|IFNA)\s*\(/i.test(cell.raw) &&
                cell.raw.replace(/"(?:[^"]|"")*"/g, '""').includes(":")) ||
              /[{}#]/.test(cell.raw.replace(/"(?:[^"]|"")*"/g, '""')) ||
              /^=\s*[^()+*/]*:/.test(cell.raw) ||
              (workbook.definedNames?.length &&
                /^=\s*[A-Za-z_][A-Za-z0-9_.]*\s*$/.test(cell.raw))),
        )
        .map(([formulaAddress]) => formulaAddress)
        .sort((leftAddress, rightAddress) => {
          const leftPosition = _t(leftAddress),
            rightPosition = _t(rightAddress);
          return (
            leftPosition.row - rightPosition.row ||
            leftPosition.col - rightPosition.col
          );
        }),
    );
  const prepareFormulaText = (
      formulaText,
      formulaSheetId,
      address,
      shouldResolveNames = true,
    ) => {
      const checkFormulaReferences = (textToCheck, we = true) => {
        const rewriteError = rewriteFormulaReferences(
          textToCheck,
          workbook,
          formulaSheetId,
          address,
          we,
        );
        if (rewriteError)
          throw (
            referenceDiagnostics.set(
              `${formulaSheetId}:${address}`,
              rewriteError,
            ),
            issue(rewriteError.error)
          );
      };
      ((formulaText = externalLinkAliases.rewrite(formulaText)),
        checkFormulaReferences(formulaText, false));
      let expandedFormula = shouldResolveNames
        ? resolveDefinedNameValue(
            formulaText,
            workbook,
            formulaSheetId,
            [],
            accessPolicy?.allowName,
          )
        : formulaText;
      ((expandedFormula = externalLinkAliases.rewrite(expandedFormula)),
        checkFormulaReferences(expandedFormula));
      let rewrittenText = "",
        scanPosition = 0;
      for (; scanPosition < expandedFormula.length;) {
        const sheetSpillMatch =
          /^(?:'((?:[^']|'')*)'|([A-Za-z_][A-Za-z0-9_.]*))!\s*(\$?[A-Za-z]{1,3}\$?[1-9]\d*)#/i.exec(
            expandedFormula.slice(scanPosition),
          );
        if (sheetSpillMatch) {
          const referencedSheet = sheetsByName.get(
            (
              sheetSpillMatch[1]?.replace(/''/g, "'") ?? sheetSpillMatch[2]
            ).toLowerCase(),
          );
          if (!referencedSheet) throw issue("#REF!");
          const anchorPosition = _t(sheetSpillMatch[3]),
            anchorAddress = `${columnIndexToLetters(anchorPosition.col)}${anchorPosition.row + 1}`,
            cell = readCell(referencedSheet.id, anchorAddress),
            spillRange = spillRanges.get(
              `${referencedSheet.id}:${anchorAddress}`,
            );
          if (!spillRange) throw isFormulaError(cell) ? cell : issue("#REF!");
          ((rewrittenText += `'${referencedSheet.name.replace(/'/g, "''")}'!${spillRange}`),
            (scanPosition += sheetSpillMatch[0].length));
          continue;
        }
        if (
          expandedFormula[scanPosition] === '"' ||
          expandedFormula[scanPosition] === "'"
        ) {
          const quoteChar = expandedFormula[scanPosition],
            quoteStart = scanPosition++;
          for (; scanPosition < expandedFormula.length;)
            if (expandedFormula[scanPosition++] === quoteChar)
              if (expandedFormula[scanPosition] === quoteChar) scanPosition++;
              else break;
          rewrittenText += expandedFormula.slice(quoteStart, scanPosition);
          continue;
        }
        const leadingCommaMatch =
          /^(SEQUENCE|FILTER|SORT|SORTBY|UNIQUE|TRANSPOSE|TAKE|DROP|TEXTSPLIT)\s*\(\s*,/i.exec(
            expandedFormula.slice(scanPosition),
          );
        if (leadingCommaMatch) {
          ((rewrittenText += `${leadingCommaMatch[1]}(${leadingCommaMatch[1].toUpperCase() === "SEQUENCE" ? "1" : "#VALUE!"},`),
            (scanPosition += leadingCommaMatch[0].length));
          continue;
        }
        const structuredRefMatch = /^(?:[A-Za-z_][A-Za-z0-9_.]*)?\[/.exec(
          expandedFormula.slice(scanPosition),
        );
        if (structuredRefMatch) {
          const structuredRefStart = scanPosition;
          scanPosition += structuredRefMatch[0].length;
          let bracketDepth = 1;
          for (; scanPosition < expandedFormula.length && bracketDepth;)
            (expandedFormula[scanPosition] === "[" && bracketDepth++,
              expandedFormula[scanPosition] === "]" && bracketDepth--,
              scanPosition++);
          const resolvedStructuredRef =
            !bracketDepth &&
            rewriteStructuredReferences(
              workbook,
              formulaSheetId,
              expandedFormula.slice(structuredRefStart, scanPosition),
              address,
            );
          if (!resolvedStructuredRef) throw issue("#REF!");
          const structuredRefSheet = sheetsById.get(
            resolvedStructuredRef.sheetId,
          );
          rewrittenText += `'${structuredRefSheet.name.replace(/'/g, "''")}'!${resolvedStructuredRef.range}`;
          continue;
        }
        const spillRefMatch = /^(\$?[A-Za-z]{1,3}\$?[1-9]\d*)#/.exec(
          expandedFormula.slice(scanPosition),
        );
        if (spillRefMatch) {
          const spillRefPosition = _t(spillRefMatch[1]),
            spillRefAddress = `${columnIndexToLetters(spillRefPosition.col)}${spillRefPosition.row + 1}`,
            cell = readCell(formulaSheetId, spillRefAddress),
            spillRefRange = spillRanges.get(
              `${formulaSheetId}:${spillRefAddress}`,
            );
          if (!spillRefRange)
            throw isFormulaError(cell) ? cell : issue("#REF!");
          ((rewrittenText += spillRefRange),
            (scanPosition += spillRefMatch[0].length));
          continue;
        }
        rewrittenText += expandedFormula[scanPosition++];
      }
      return rewrittenText.replace(/\b_xlfn\.(?:_xlws\.)?/gi, "");
    },
    getSheetByName = (sheetName) => {
      const matchedSheet = sheetsByName.get((sheetName ?? "").toLowerCase());
      if (!matchedSheet) throw issue("#REF!");
      return matchedSheet;
    },
    createFormulaEngine = () => {
      const workbookFormulaEngine = createWorkbookFormulaEngine(
        new FormulaParser({
          onVariable: () => null,
          onCell: ({ sheet, row, col }) => {
            const targetSheet = getSheetByName(sheet);
            if (
              !Number.isInteger(row) ||
              !Number.isInteger(col) ||
              row < 1 ||
              col < 1 ||
              row > targetSheet.rowCount ||
              col > targetSheet.colCount
            )
              throw issue("#REF!");
            return readCell(
              targetSheet.id,
              `${columnIndexToLetters(col - 1)}${row}`,
            );
          },
          onRange: ({ sheet, from, to: rangeEnd }) => {
            const rangeSheet = getSheetByName(sheet),
              row = rangeEnd.row,
              col = rangeEnd.col;
            if (
              from.row < 1 ||
              from.col < 1 ||
              row > rangeSheet.rowCount ||
              col > rangeSheet.colCount ||
              row < from.row ||
              col < from.col
            )
              throw issue("#REF!");
            if ((row - from.row + 1) * (col - from.col + 1) > 2e4)
              throw issue("#LIMIT!");
            const reserveIssue = budgetTracker.reserve(
              (row - from.row + 1) * (32 + (col - from.col + 1) * 8),
            );
            if (reserveIssue) throw issue(reserveIssue);
            const rangeRows = [];
            for (let rangeRow = from.row; rangeRow <= row; rangeRow++) {
              const rangeCells = [];
              for (
                let rangeColumn = from.col;
                rangeColumn <= col;
                rangeColumn++
              ) {
                const cell = readCell(
                  rangeSheet.id,
                  `${columnIndexToLetters(rangeColumn - 1)}${rangeRow}`,
                );
                if (
                  isFormulaError(cell) &&
                  ["#LIMIT!", "#TIMEOUT!", "#BLOCKED!"].includes(cell.error)
                )
                  throw cell;
                rangeCells.push(cell);
              }
              rangeRows.push(rangeCells);
            }
            return rangeRows;
          },
        }),
        (le, ce, formulaContext) =>
          filteredRowContext(le, ce, {
            currentSheetName: formulaContext.position?.sheet,
            retrieve: (reference) => formulaContext.retrieveRef(reference),
            issue: issue,
            isIssue: isFormulaError,
            readCell: readCell,
            readMetadata:
              accessPolicy || calculationContext
                ? (sheetId, address) => {
                    if (
                      (calculationContext?.read(`${sheetId}:${address}`),
                      accessPolicy && !accessPolicy.allowRead(sheetId, address))
                    )
                      throw issue("#BLOCKED!");
                    accessPolicy?.step({
                      kind: "metadata",
                      sheetId: sheetId,
                      address: address,
                      operation: "Inspect cell metadata",
                      result: null,
                    });
                  }
                : undefined,
            isActive: (checkSheetId, checkAddress) =>
              activeFormulaCells.has(`${checkSheetId}:${checkAddress}`),
            expandedFormula: (cell, nameScope) => {
              try {
                return resolveDefinedNameValue(
                  String(cell.raw),
                  workbook,
                  nameScope,
                );
              } catch {
                return String(cell.raw);
              }
            },
          }),
        issue,
        clock,
        (callInfo) => {
          const position = workbookFormulaEngine.position;
          if (!position) return;
          const sheetId = getSheetByName(position.sheet).id,
            address = `${columnIndexToLetters(position.col - 1)}${position.row}`,
            diagnosticKey = `${sheetId}:${address}:${callInfo.functionName}`;
          if (
            financialDiagnostics.size >= 1e4 &&
            !financialDiagnostics.has(diagnosticKey)
          )
            throw issue("#LIMIT!");
          financialDiagnostics.set(diagnosticKey, {
            ...callInfo,
            sheetId: sheetId,
            address: address,
          });
        },
        (refValue, rowOffset, colOffset) => {
          const refDescriptor =
            refValue && typeof refValue == "object" && "ref" in refValue
              ? refValue.ref
              : undefined;
          if (!refDescriptor?.from || !refDescriptor.to) return;
          const refSheet = getSheetByName(
              refDescriptor.sheet ?? workbookFormulaEngine.position?.sheet,
            ),
            refAddress = `${columnIndexToLetters(refDescriptor.from.col + colOffset - 1)}${refDescriptor.from.row + rowOffset}`;
          if (accessPolicy && !accessPolicy.allowRead(refSheet.id, refAddress))
            throw issue("#BLOCKED!");
          calculationContext?.read(`${refSheet.id}:${refAddress}`);
          const cell = refSheet.cells[refAddress];
          if (
            cell?.literal ||
            typeof cell?.raw != "string" ||
            !cell.raw.startsWith("=")
          )
            return;
          const raw = cell.raw;
          return (shiftAmount) => {
            if (criterionDepth >= 32) throw issue("#LIMIT!");
            const criterionEngineIndex = criterionDepth++;
            try {
              if (
                accessPolicy?.allowDatabaseCriterion &&
                !accessPolicy.allowDatabaseCriterion(
                  refSheet.id,
                  shiftFormulaReferences(raw, shiftAmount, 0),
                )
              )
                throw issue("#BLOCKED!");
              const definedNameValue = resolveDefinedNameValue(
                  raw,
                  workbook,
                  refSheet.id,
                  [],
                  accessPolicy?.allowName,
                ),
                shiftedFormula = shiftFormulaReferences(
                  definedNameValue,
                  shiftAmount,
                  0,
                );
              if (shiftedFormula.length > fileImportLimits.maxFormulaLength)
                throw issue("#LIMIT!");
              return isReusableFormula(shiftedFormula)
                ? evaluateReusable(
                    shiftedFormula.slice(1),
                    refSheet.id,
                    refAddress,
                  )
                : (criterionFormulaEngines[criterionEngineIndex] ??=
                    createFormulaEngine()).parse(
                    prepareFormulaText(
                      shiftedFormula.slice(1),
                      refSheet.id,
                      refAddress,
                    ),
                    {
                      sheet: refSheet.name,
                      row: refDescriptor.from.row + rowOffset,
                      col: refDescriptor.from.col + colOffset,
                    },
                    true,
                  );
            } finally {
              criterionDepth--;
            }
          };
        },
      );
      if (installedExtensionFunctions.length) {
        const originalCallFunction = workbookFormulaEngine._callFunction.bind(
          workbookFormulaEngine,
        );
        workbookFormulaEngine._callFunction = (functionName, functionArgs) => {
          const upperFunctionName = functionName
            .replace(/^_xlfn\./i, "")
            .toUpperCase();
          if (
            !installedExtensionFunctions.some(
              (installedExtensionFunction) =>
                installedExtensionFunction.name === upperFunctionName,
            )
          )
            return originalCallFunction(functionName, functionArgs);
          const position = workbookFormulaEngine.position;
          if (!position) return issue("#NAME?");
          const callSheetId = getSheetByName(position.sheet).id,
            callAddress = `${columnIndexToLetters(position.col - 1)}${position.row}`,
            resolvedArgs = functionArgs.map((argument) =>
              workbookFormulaEngine.retrieveRef(argument),
            ),
            formulaErrorArg = findFormulaErrorArg(resolvedArgs);
          if (formulaErrorArg) return formulaErrorArg;
          const extensionFunctionCall = resolveExtensionFunctionCall(
            workbook,
            upperFunctionName,
            resolvedArgs,
            callSheetId,
            callAddress,
          );
          if (extensionFunctionCall.error)
            return issue(extensionFunctionCall.error);
          const value = extensionFunctionCall.value;
          return value && !Array.isArray(value) && typeof value == "object"
            ? value.label
            : value;
        };
      }
      const runWithBudget = (computeValue) => {
          const budgetIssue =
            budgetTracker.work() ?? calculationContext?.check();
          if (budgetIssue) return issue(budgetIssue);
          const computedResult = computeValue(),
            valueIssue =
              budgetTracker.value(computedResult) ??
              calculationContext?.check();
          return valueIssue ? issue(valueIssue) : computedResult;
        },
        boundCallFunction = workbookFormulaEngine._callFunction.bind(
          workbookFormulaEngine,
        ),
        boundApplyInfix = workbookFormulaEngine.utils._applyInfix.bind(
          workbookFormulaEngine.utils,
        ),
        boundApplyPrefix = workbookFormulaEngine.utils._applyPrefix.bind(
          workbookFormulaEngine.utils,
        ),
        boundApplyPostfix = workbookFormulaEngine.utils._applyPostfix.bind(
          workbookFormulaEngine.utils,
        );
      if (
        ((workbookFormulaEngine._callFunction = (callName, callArgs) =>
          runWithBudget(() => boundCallFunction(callName, callArgs))),
        (workbookFormulaEngine.utils._applyInfix = (
          leftOperand,
          infixOperator,
          rightOperand,
        ) =>
          runWithBudget(() =>
            boundApplyInfix(leftOperand, infixOperator, rightOperand),
          )),
        (workbookFormulaEngine.utils._applyPrefix = (
          prefixOperators,
          prefixOperand,
          ge,
        ) =>
          runWithBudget(() =>
            boundApplyPrefix(prefixOperators, prefixOperand, ge),
          )),
        (workbookFormulaEngine.utils._applyPostfix = (
          postfixOperand,
          ce,
          postfixOperator,
        ) =>
          runWithBudget(() =>
            boundApplyPostfix(postfixOperand, ce, postfixOperator),
          )),
        accessPolicy)
      ) {
        const recordStep = (stepKind, operation, inputs, computeStep) => {
            const position = workbookFormulaEngine.position,
              sheetId = getSheetByName(position.sheet).id,
              address = `${columnIndexToLetters(position.col - 1)}${position.row}`;
            try {
              const stepResult = computeStep();
              return (
                accessPolicy.step({
                  kind: stepKind,
                  sheetId: sheetId,
                  address: address,
                  operation: operation,
                  inputs: inputs,
                  result: stepResult,
                }),
                stepResult
              );
            } catch (error) {
              throw (
                accessPolicy.step({
                  kind: stepKind,
                  sheetId: sheetId,
                  address: address,
                  operation: operation,
                  inputs: inputs,
                  result: isFormulaError(error) ? error : issue("#ERROR!"),
                }),
                error
              );
            }
          },
          baseCallFunction = workbookFormulaEngine._callFunction.bind(
            workbookFormulaEngine,
          ),
          baseApplyInfix = workbookFormulaEngine.utils._applyInfix.bind(
            workbookFormulaEngine.utils,
          ),
          baseApplyPrefix = workbookFormulaEngine.utils._applyPrefix.bind(
            workbookFormulaEngine.utils,
          ),
          baseApplyPostfix = workbookFormulaEngine.utils._applyPostfix.bind(
            workbookFormulaEngine.utils,
          );
        ((workbookFormulaEngine._callFunction = (traceName, traceArgs) =>
          recordStep("function", traceName, traceArgs, () =>
            baseCallFunction(traceName, traceArgs),
          )),
          (workbookFormulaEngine.utils._applyInfix = (
            tracedLeft,
            tracedOperator,
            tracedRight,
          ) =>
            recordStep(
              "operator",
              tracedOperator,
              [tracedLeft.val, tracedRight.val],
              () => baseApplyInfix(tracedLeft, tracedOperator, tracedRight),
            )),
          (workbookFormulaEngine.utils._applyPrefix = (
            tracedPrefixOperators,
            tracedPrefixOperand,
            ye,
          ) =>
            recordStep(
              "operator",
              tracedPrefixOperators.join(""),
              [tracedPrefixOperand],
              () =>
                baseApplyPrefix(tracedPrefixOperators, tracedPrefixOperand, ye),
            )),
          (workbookFormulaEngine.utils._applyPostfix = (
            tracedPostfixOperand,
            Ae,
            tracedPostfixOperator,
          ) =>
            recordStep(
              "operator",
              tracedPostfixOperator,
              [tracedPostfixOperand],
              () =>
                baseApplyPostfix(
                  tracedPostfixOperand,
                  Ae,
                  tracedPostfixOperator,
                ),
            )));
      }
      return workbookFormulaEngine;
    },
    evaluateReusable = (reusableFormulaText, sheetId, address) =>
      evaluateReusableFormula(reusableFormulaText, workbook, sheetId, {
        supported: supported,
        issue: issue,
        isIssue: isFormulaError,
        allowName: accessPolicy?.allowName,
        step: accessPolicy
          ? (traceKind, operation, traceResult) =>
              accessPolicy.step({
                kind: traceKind,
                sheetId: sheetId,
                address: address,
                operation: operation,
                result: traceResult,
              })
          : undefined,
        diagnostic: (diagnosticCode, diagnosticMessage) =>
          namedFormulaDiagnostics.set(`${sheetId}:${address}`, {
            sheetId: sheetId,
            address: address,
            error: diagnosticCode,
            message: diagnosticMessage,
          }),
        parse: (parseText, parseSheetId, lambdaArgs) => {
          const rewrittenFormula = prepareFormulaText(
              parseText,
              parseSheetId,
              address,
              false,
            ),
            textWithoutStrings = rewrittenFormula.replace(
              /"(?:[^"]|"")*"/g,
              '""',
            );
          if (
            /[\[\]|]/.test(textWithoutStrings) ||
            /(?:https?:|file:|\\\\)/i.test(textWithoutStrings)
          )
            throw issue("#BLOCKED!");
          if (
            Array.from(
              textWithoutStrings
                .replace(/'(?:[^']|'')*'(?=\s*!)/g, "Sheet")
                .matchAll(/([A-Za-z_][A-Za-z0-9_.]*)\s*\(/g),
            ).some(
              (functionMatch) =>
                functionMatch[1] !== y5 &&
                !supported.has(functionMatch[1].toUpperCase()),
            )
          )
            throw issue("#NAME?");
          const engineIndex = reusableFormulaDepth++,
            reusableEngine = (reusableFormulaEngines[engineIndex] ??=
              createFormulaEngine()),
            previousCallFunction = reusableEngine._callFunction;
          reusableEngine._callFunction = (lambdaCallName, lambdaCallArgs) =>
            lambdaCallName === y5
              ? lambdaCallArgs.length === 1 &&
                Number.isInteger(lambdaCallArgs[0]) &&
                Number(lambdaCallArgs[0]) >= 0 &&
                Number(lambdaCallArgs[0]) < lambdaArgs.length
                ? lambdaArgs[Number(lambdaCallArgs[0])]
                : issue("#NAME?")
              : previousCallFunction(lambdaCallName, lambdaCallArgs);
          try {
            const parsePosition = _t(address);
            return reusableEngine.parse(
              rewrittenFormula,
              {
                sheet: sheetsById.get(parseSheetId).name,
                row: parsePosition.row + 1,
                col: parsePosition.col + 1,
              },
              true,
            );
          } finally {
            ((reusableEngine._callFunction = previousCallFunction),
              reusableFormulaDepth--);
          }
        },
      }),
    readCell = (sheetId, address) => {
      const readBudgetIssue =
        budgetTracker.work() ?? calculationContext?.check();
      if (readBudgetIssue) return issue(readBudgetIssue);
      if (!accessPolicy) return resolveCellValue(sheetId, address);
      if (!accessPolicy.allowRead(sheetId, address)) return issue("#BLOCKED!");
      const cellValue = resolveCellValue(sheetId, address);
      return (
        accessPolicy.step({
          kind: "read",
          sheetId: sheetId,
          address: address,
          operation: "Read cell",
          result: cellValue,
        }),
        cellValue
      );
    },
    resolveCellValue = (sheetId, cellAddress) => {
      if (++cellReadCount > workbookLimits.maxCells) return issue("#LIMIT!");
      let firstRow, firstCol;
      try {
        ({ row: firstRow, col: firstCol } = _t(cellAddress));
      } catch {
        return issue("#REF!");
      }
      const sheet = sheetsById.get(sheetId);
      if (!sheet || firstRow >= sheet.rowCount || firstCol >= sheet.colCount)
        return issue("#REF!");
      cellAddress = `${columnIndexToLetters(firstCol)}${firstRow + 1}`;
      const cellKey = `${sheetId}:${cellAddress}`,
        externalValue = externalLinkAliases.values.get(cellKey);
      if (externalValue)
        return externalValue.type === "error"
          ? issue(String(externalValue.value))
          : externalValue.value;
      calculationContext?.read(cellKey);
      const contextValue = calculationContext?.value(
        cellKey,
        activeFormulaCells.size > 0 || parsingCells.size > 0,
      );
      if (contextValue)
        return contextValue.type === "error"
          ? issue(String(contextValue.value))
          : contextValue.value;
      for (const activeFormulaCell of activeFormulaCells)
        readDependencies.get(activeFormulaCell)?.add(cellKey);
      if (activeFormulaCells.has(cellKey) || parsingCells.has(cellKey))
        return (
          calculationContext?.cycle([
            ...activeFormulaCells,
            ...parsingCells,
            cellKey,
          ]),
          issue("#CYCLE!")
        );
      if (computedCellValues.has(cellKey))
        return computedCellValues.get(cellKey);
      if (spillCellValues.has(cellKey))
        return spillCellValues.get(cellKey).value;
      const cell = sheet.cells[cellAddress];
      if (!cell || cell.raw === null) {
        const spillCandidates = spillCandidateAddresses.get(sheetId) ?? [];
        if (spillCandidates.length > 1e3) return issue("#LIMIT!");
        for (const candidateAddress of spillCandidates) {
          if (candidateAddress === cellAddress) continue;
          const candidatePosition = _t(candidateAddress);
          if (
            !(
              candidatePosition.row > firstRow ||
              candidatePosition.col > firstCol ||
              (firstRow - candidatePosition.row + 1) *
                (firstCol - candidatePosition.col + 1) >
                fileImportLimits.maxSpillCells
            ) &&
            !(
              activeFormulaCells.has(`${sheetId}:${candidateAddress}`) ||
              parsingCells.has(`${sheetId}:${candidateAddress}`)
            ) &&
            (readCell(sheetId, candidateAddress), spillCellValues.has(cellKey))
          )
            return spillCellValues.get(cellKey).value;
        }
        return cell?.raw ?? null;
      }
      if (
        typeof cell.raw != "string" ||
        !cell.raw.startsWith("=") ||
        cell.literal
      )
        return cell.raw;
      const raw = cell.raw;
      calculationContext?.enter(cellKey);
      let formulaResult = (() => {
        if (
          activeFormulaCells.size + parsingCells.size >=
            fileImportLimits.maxFormulaDepth ||
          raw.length > fileImportLimits.maxFormulaLength ||
          !isWithinNestingLimit(raw)
        )
          return issue("#LIMIT!");
        let formulaBody;
        const isReusable = isReusableFormula(raw);
        parsingCells.add(cellKey);
        try {
          if (
            ((formulaBody = isReusable
              ? raw.slice(1)
              : prepareFormulaText(raw.slice(1), sheetId, cellAddress)),
            !isWithinNestingLimit(formulaBody))
          )
            return issue("#LIMIT!");
        } catch (error) {
          return isFormulaError(error)
            ? issue(error.error)
            : issue(
                error instanceof Error && /^#/.test(error.message)
                  ? error.message
                  : "#NAME?",
              );
        } finally {
          parsingCells.delete(cellKey);
        }
        const bodyWithoutStrings = formulaBody.replace(/"(?:[^"]|"")*"/g, '""');
        if (
          !isReusable &&
          (/[\[\]|]/.test(bodyWithoutStrings) ||
            /(?:https?:|file:|\\\\)/i.test(bodyWithoutStrings))
        )
          return issue("#BLOCKED!");
        if (
          !isReusable &&
          Array.from(
            bodyWithoutStrings
              .replace(/'(?:[^']|'')*'(?=\s*!)/g, "Sheet")
              .matchAll(/([A-Za-z_][A-Za-z0-9_.]*)\s*\(/g),
          ).some((callMatch) => !supported.has(callMatch[1].toUpperCase()))
        )
          return (
            namedFormulaDiagnostics.set(cellKey, {
              sheetId: sheetId,
              address: cellAddress,
              error: "#NAME?",
              message:
                "The formula calls an unknown or unsupported custom function. Imported VBA, add-in and external custom functions cannot execute.",
            }),
            issue("#NAME?")
          );
        const size = activeFormulaCells.size;
        (activeFormulaCells.add(cellKey),
          readDependencies.set(cellKey, new Set()));
        let computedValue;
        try {
          const nestedEngine = (nestedFormulaEngines[size] ??=
              createFormulaEngine()),
            rawResult = isReusable
              ? evaluateReusable(formulaBody, sheetId, cellAddress)
              : nestedEngine.parse(
                  formulaBody,
                  { sheet: sheet.name, row: firstRow + 1, col: firstCol + 1 },
                  true,
                );
          if (isFormulaError(rawResult)) computedValue = issue(rawResult.error);
          else if (typeof rawResult == "number")
            computedValue = Number.isFinite(rawResult)
              ? rawResult
              : issue("#NUM!");
          else if (rawResult === null) computedValue = 0;
          else if (typeof rawResult == "boolean") computedValue = rawResult;
          else if (typeof rawResult == "string")
            computedValue =
              rawResult.length <= 32767 ? rawResult : issue("#LIMIT!");
          else if (Array.isArray(rawResult)) {
            const rowMatrix = rawResult.every(Array.isArray)
                ? rawResult
                : [rawResult],
              spillColumnCount = rowMatrix[0]?.length ?? 0;
            if (
              !rowMatrix.length ||
              !spillColumnCount ||
              rowMatrix.some(
                (rowValues) => rowValues.length !== spillColumnCount,
              )
            )
              computedValue = issue("#CALC!");
            else if (
              rowMatrix.length * spillColumnCount >
                fileImportLimits.maxSpillCells ||
              spilledCellCount + rowMatrix.length * spillColumnCount >
                workbookLimits.maxCells
            )
              computedValue = issue("#LIMIT!");
            else if (
              firstRow + rowMatrix.length > sheet.rowCount ||
              firstCol + spillColumnCount > sheet.colCount
            )
              (spillDiagnostics.set(cellKey, {
                sheetId: sheetId,
                anchor: cellAddress,
                reason: "bounds",
                message:
                  "The array extends beyond the worksheet edge. Move the formula or reduce its output.",
              }),
                (computedValue = issue("#SPILL!")));
            else {
              let hasObstruction = false,
                hasCycle = false;
              const spillBounds = {
                  firstRow: firstRow,
                  lastRow: firstRow + rowMatrix.length - 1,
                  firstCol: firstCol,
                  lastCol: firstCol + spillColumnCount - 1,
                },
                overlapsSpill = (rangeText) => {
                  const rangeAddress = parseRangeAddress(rangeText);
                  return (
                    rangeAddress.firstRow <= spillBounds.lastRow &&
                    rangeAddress.lastRow >= spillBounds.firstRow &&
                    rangeAddress.firstCol <= spillBounds.lastCol &&
                    rangeAddress.lastCol >= spillBounds.firstCol
                  );
                },
                overlappingMerge = sheet.merges?.find(overlapsSpill),
                overlappingTable = sheet.tables?.find((table) =>
                  overlapsSpill(table.range),
                );
              if (overlappingMerge || overlappingTable) {
                hasObstruction = true;
                const obstruction = overlappingMerge ?? overlappingTable.range;
                spillDiagnostics.set(cellKey, {
                  sheetId: sheetId,
                  anchor: cellAddress,
                  reason: overlappingMerge ? "merged" : "table",
                  obstruction: obstruction,
                  message: overlappingMerge
                    ? `The array overlaps merged cells ${obstruction}. Unmerge those cells or move the formula.`
                    : `The array overlaps table ${overlappingTable.name} (${obstruction}). Place the formula outside the table.`,
                });
              }
              for (
                let obstructionRow = 0;
                obstructionRow < rowMatrix.length && !hasObstruction;
                obstructionRow++
              )
                for (
                  let obstructionColumn = 0;
                  obstructionColumn < spillColumnCount;
                  obstructionColumn++
                ) {
                  const obstruction = `${columnIndexToLetters(firstCol + obstructionColumn)}${firstRow + obstructionRow + 1}`,
                    cell = sheet.cells[obstruction];
                  if (
                    readDependencies
                      .get(cellKey)
                      ?.has(`${sheetId}:${obstruction}`)
                  ) {
                    hasCycle = true;
                    break;
                  }
                  if (
                    obstruction !== cellAddress &&
                    ((cell && cell.raw !== null) ||
                      spillCellValues.has(`${sheetId}:${obstruction}`))
                  ) {
                    ((hasObstruction = true),
                      spillDiagnostics.set(cellKey, {
                        sheetId: sheetId,
                        anchor: cellAddress,
                        reason: "occupied",
                        obstruction: obstruction,
                        message: `Cell ${obstruction} blocks the array. Move the formula or review and clear the obstruction; existing evidence was preserved.`,
                      }));
                    break;
                  }
                }
              if (hasCycle) computedValue = issue("#CYCLE!");
              else if (hasObstruction) computedValue = issue("#SPILL!");
              else {
                const normalizeSpillValue = (rawSpillValue) =>
                  isFormulaError(rawSpillValue)
                    ? rawSpillValue
                    : rawSpillValue === null
                      ? 0
                      : typeof rawSpillValue == "number"
                        ? Number.isFinite(rawSpillValue)
                          ? rawSpillValue
                          : issue("#NUM!")
                        : typeof rawSpillValue == "boolean"
                          ? rawSpillValue
                          : typeof rawSpillValue == "string"
                            ? rawSpillValue.length <= 32767
                              ? rawSpillValue
                              : issue("#LIMIT!")
                            : issue("#VALUE!");
                for (
                  let storeRowIndex = 0;
                  storeRowIndex < rowMatrix.length;
                  storeRowIndex++
                )
                  for (
                    let storeColumnIndex = 0;
                    storeColumnIndex < spillColumnCount;
                    storeColumnIndex++
                  ) {
                    const storeAddress = `${columnIndexToLetters(firstCol + storeColumnIndex)}${firstRow + storeRowIndex + 1}`;
                    storeAddress !== cellAddress &&
                      spillCellValues.set(`${sheetId}:${storeAddress}`, {
                        value: normalizeSpillValue(
                          rowMatrix[storeRowIndex][storeColumnIndex],
                        ),
                        anchor: cellAddress,
                      });
                  }
                (spillRanges.set(
                  cellKey,
                  `${cellAddress}:${columnIndexToLetters(firstCol + spillColumnCount - 1)}${firstRow + rowMatrix.length}`,
                ),
                  (spilledCellCount += rowMatrix.length * spillColumnCount),
                  (computedValue = normalizeSpillValue(rowMatrix[0][0])));
              }
            }
          } else computedValue = issue("#VALUE!");
        } catch (error) {
          const errorDetails =
            error && typeof error == "object" && "details" in error
              ? error.details
              : undefined;
          computedValue = isFormulaError(errorDetails)
            ? issue(errorDetails.error)
            : isFormulaError(error)
              ? issue(error.error)
              : issue("#ERROR!");
        } finally {
          (activeFormulaCells.delete(cellKey),
            readDependencies.delete(cellKey));
        }
        return computedValue;
      })();
      const finalIssue =
        budgetTracker.value(formulaResult) ?? calculationContext?.check();
      return (
        finalIssue && (formulaResult = issue(finalIssue)),
        computedCellValues.set(cellKey, formulaResult),
        calculationContext?.leave(
          cellKey,
          isFormulaError(formulaResult)
            ? { type: "error", value: formulaResult.error }
            : {
                type:
                  formulaResult === null
                    ? "blank"
                    : typeof formulaResult == "string"
                      ? "text"
                      : typeof formulaResult == "boolean"
                        ? "boolean"
                        : "number",
                value: formulaResult,
              },
        ),
        accessPolicy?.step({
          kind: "formula",
          sheetId: sheetId,
          address: cellAddress,
          operation: cell.raw,
          result: formulaResult,
        }),
        formulaResult
      );
    },
    getCellValue = (targetSheetId, targetAddress) =>
      getCellValue.getTypedValue(targetSheetId, targetAddress).value;
  getCellValue.getTypedValue = (typedSheetId, typedAddress) => {
    if (calculationContext?.running?.() === false)
      try {
        const typedPosition = _t(typedAddress),
          sheet = sheetsById.get(typedSheetId);
        if (
          !sheet ||
          typedPosition.row >= sheet.rowCount ||
          typedPosition.col >= sheet.colCount
        )
          return { type: "error", value: "#REF!" };
        const typedCellAddress = `${columnIndexToLetters(typedPosition.col)}${typedPosition.row + 1}`,
          typedCellKey = `${typedSheetId}:${typedCellAddress}`,
          cachedValue = calculationContext.value(typedCellKey, false);
        if (cachedValue) return { ...cachedValue };
        const typedValue =
          computedCellValues.get(typedCellKey) ??
          spillCellValues.get(typedCellKey)?.value ??
          sheet.cells[typedCellAddress]?.raw ??
          null;
        return isFormulaError(typedValue)
          ? { type: "error", value: typedValue.error }
          : {
              type:
                typedValue === null
                  ? "blank"
                  : typeof typedValue == "string"
                    ? "text"
                    : typeof typedValue == "boolean"
                      ? "boolean"
                      : "number",
              value: typedValue,
            };
      } catch {
        return { type: "error", value: "#REF!" };
      }
    budgetTracker.begin();
    try {
      cellReadCount = 0;
      const cell = readCell(typedSheetId, typedAddress);
      return isFormulaError(cell)
        ? { type: "error", value: cell.error }
        : {
            type:
              cell === null
                ? "blank"
                : typeof cell == "string"
                  ? "text"
                  : typeof cell == "boolean"
                    ? "boolean"
                    : "number",
            value: cell,
          };
    } finally {
      budgetTracker.end();
    }
  };
  let sharedFormulaEngine;
  return (
    (getCellValue.evaluateFormula = (evalSheetId, evalAddress, evalFormula) => {
      (budgetTracker.begin(),
        (cellReadCount = 0),
        referenceDiagnostics.delete(`${evalSheetId}:${evalAddress}`),
        namedFormulaDiagnostics.delete(`${evalSheetId}:${evalAddress}`));
      const expressionKey = `expression:${evalSheetId}:${evalAddress}`;
      try {
        const evalSheet = sheetsById.get(evalSheetId),
          evalPosition = _t(evalAddress);
        if (
          !evalSheet ||
          evalPosition.row >= evalSheet.rowCount ||
          evalPosition.col >= evalSheet.colCount
        )
          return "#REF!";
        if (typeof evalFormula != "string" || !evalFormula.startsWith("="))
          return "#VALUE!";
        if (
          evalFormula.length > fileImportLimits.maxFormulaLength ||
          !isWithinNestingLimit(evalFormula)
        )
          return issue("#LIMIT!").error;
        const evalIsReusable = isReusableFormula(evalFormula),
          evalFormulaBody = evalIsReusable
            ? evalFormula.slice(1)
            : prepareFormulaText(
                evalFormula.slice(1),
                evalSheetId,
                evalAddress,
              );
        if (!isWithinNestingLimit(evalFormulaBody))
          return issue("#LIMIT!").error;
        const evalBodyWithoutStrings = evalFormulaBody.replace(
          /"(?:[^"]|"")*"/g,
          '""',
        );
        if (
          !evalIsReusable &&
          (/[\[\]|]/.test(evalBodyWithoutStrings) ||
            /(?:https?:|file:|\\\\)/i.test(evalBodyWithoutStrings))
        )
          return issue("#BLOCKED!").error;
        if (
          !evalIsReusable &&
          Array.from(
            evalBodyWithoutStrings
              .replace(/'(?:[^']|'')*'(?=\s*!)/g, "Sheet")
              .matchAll(/([A-Za-z_][A-Za-z0-9_.]*)\s*\(/g),
          ).some(
            (evalFunctionMatch) =>
              !supported.has(evalFunctionMatch[1].toUpperCase()),
          )
        )
          return "#NAME?";
        (activeFormulaCells.add(expressionKey),
          readDependencies.set(expressionKey, new Set()));
        const evalResult = evalIsReusable
            ? evaluateReusable(evalFormulaBody, evalSheetId, evalAddress)
            : (sharedFormulaEngine ??= createFormulaEngine()).parse(
                evalFormulaBody,
                {
                  sheet: evalSheet.name,
                  row: evalPosition.row + 1,
                  col: evalPosition.col + 1,
                },
                true,
              ),
          evalIssue =
            budgetTracker.value(evalResult) ?? calculationContext?.check();
        return evalIssue
          ? issue(evalIssue).error
          : isFormulaError(evalResult)
            ? issue(evalResult.error).error
            : Array.isArray(evalResult)
              ? "#VALUE!"
              : evalResult === null
                ? 0
                : typeof evalResult == "number"
                  ? Number.isFinite(evalResult)
                    ? evalResult
                    : "#NUM!"
                  : typeof evalResult == "boolean"
                    ? evalResult
                    : typeof evalResult == "string"
                      ? evalResult.length <= 32767
                        ? evalResult
                        : issue("#LIMIT!").error
                      : "#VALUE!";
      } catch (error) {
        const evalErrorDetails =
          error && typeof error == "object" && "details" in error
            ? error.details
            : undefined;
        return isFormulaError(evalErrorDetails)
          ? issue(evalErrorDetails.error).error
          : isFormulaError(error)
            ? issue(error.error).error
            : error instanceof Error && /^#/.test(error.message)
              ? issue(error.message).error
              : issue("#ERROR!").error;
      } finally {
        (budgetTracker.end(),
          activeFormulaCells.delete(expressionKey),
          readDependencies.delete(expressionKey));
      }
    }),
    (getCellValue.getSpillRange = (spillSheetId, spillAddress) => (
      getCellValue(spillSheetId, spillAddress),
      spillRanges.get(`${spillSheetId}:${spillAddress}`)
    )),
    (getCellValue.getSpillDiagnostic = (
      diagnosticSheetId,
      diagnosticAddress,
    ) => {
      getCellValue(diagnosticSheetId, diagnosticAddress);
      const spillDiagnostic = spillDiagnostics.get(
        `${diagnosticSheetId}:${diagnosticAddress}`,
      );
      return spillDiagnostic ? { ...spillDiagnostic } : undefined;
    }),
    (getCellValue.getSpillAnchor = (anchorSheetId, anchorCell) => (
      getCellValue(anchorSheetId, anchorCell),
      spillCellValues.get(`${anchorSheetId}:${anchorCell}`)?.anchor
    )),
    (getCellValue.getSpillEntries = (entriesSheetId) => {
      budgetTracker.begin();
      try {
        const entryCandidates =
          spillCandidateAddresses.get(entriesSheetId) ?? [];
        if (entryCandidates.length > 1e3)
          throw (
            issue("#LIMIT!"),
            new Error("Spill discovery exceeds 1,000 possible array formulas.")
          );
        cellReadCount = 0;
        for (const scanAddress of calculationContext?.running?.() === false
          ? []
          : entryCandidates)
          if (
            (readCell(entriesSheetId, scanAddress),
            cellReadCount > workbookLimits.maxCells)
          )
            throw new Error(
              "Spill discovery exceeds 100,000 calculation reads.",
            );
        const sheetKeyPrefix = `${entriesSheetId}:`;
        return [...spillCellValues.entries()]
          .filter(([entryKey]) => entryKey.startsWith(sheetKeyPrefix))
          .map(([storedKey, { anchor }]) => ({
            address: storedKey.slice(sheetKeyPrefix.length),
            anchor,
          }));
      } finally {
        budgetTracker.end();
      }
    }),
    (getCellValue.getResourceReport = () => budgetTracker.report()),
    (getCellValue.getOperationalErrors = () => [...operationalErrors]),
    (getCellValue.getDateContext = () => ({ ...clock })),
    (getCellValue.getFinancialDiagnostics = () =>
      structuredClone([...financialDiagnostics.values()])),
    (getCellValue.peekSpill = (peekSheetId, peekAddress) => {
      const peekKey = `${peekSheetId}:${peekAddress}`,
        peekDiagnostic = spillDiagnostics.get(peekKey);
      return {
        range: spillRanges.get(peekKey),
        ...(peekDiagnostic ? { diagnostic: { ...peekDiagnostic } } : {}),
      };
    }),
    (getCellValue.getNamedFormulaDiagnostics = () =>
      [...namedFormulaDiagnostics.values()].map((namedDiagnostic) => ({
        ...namedDiagnostic,
      }))),
    (getCellValue.getReferenceDiagnostics = () =>
      [...referenceDiagnostics.values()].map((referenceDiagnostic) => ({
        ...referenceDiagnostic,
      }))),
    getCellValue
  );
}
function evaluateValidationRule(workbook) {
  const calculator = evaluateWorkbookFormulas(workbook);
  return (sheetId, address, rule) => {
    const cellValue = calculator(sheetId, address);
    let isValid = false,
      reason = "Value does not satisfy the rule.",
      formula,
      evaluationError = false;
    try {
      if (calculator.getTypedValue(sheetId, address).type === "error")
        throw Error(`Cell calculation returned ${cellValue}.`);
      if (cellValue === null || cellValue === "") isValid = rule.allowBlank;
      else if (rule.kind === "custom") {
        const cellAddress = _t(address),
          rangeAddress = parseRangeAddress(rule.range);
        formula = shiftFormulaReferences(
          rule.formula,
          cellAddress.row - rangeAddress.firstRow,
          cellAddress.col - rangeAddress.firstCol,
        );
        const formulaResult = calculator.evaluateFormula(
          sheetId,
          address,
          formula,
        );
        if (
          typeof formulaResult != "boolean" &&
          (typeof formulaResult != "number" || !Number.isFinite(formulaResult))
        )
          throw Error(
            `Custom formula must return TRUE/FALSE or a finite number${typeof formulaResult == "string" && /^#[A-Z0-9/?!]+$/.test(formulaResult) ? `; returned ${formulaResult}` : ""}.`,
          );
        ((isValid = !!formulaResult),
          (reason = "Custom formula returned FALSE or zero."));
      } else
        isValid = validateDataValidationRule(
          rule,
          cellValue,
          workbook.datePolicy,
        );
    } catch (error) {
      ((evaluationError = true),
        (reason =
          error instanceof Error
            ? error.message
            : "Validation could not be evaluated."));
    }
    if (!(isValid && !evaluationError))
      return {
        sheetId: sheetId,
        sheetName: workbook.sheets.find((sheet) => sheet.id === sheetId).name,
        address: address,
        ruleId: rule.id,
        style: evaluationError ? "stop" : (rule.errorStyle ?? "stop"),
        value: cellValue,
        title: rule.errorTitle || "Invalid value",
        message:
          rule.message ||
          "The value does not satisfy the data validation rule.",
        reason: reason,
        ...(formula ? { formula: formula } : {}),
        evaluationError: evaluationError,
      };
  };
}
function checkDataValidationBatch(book, targets, validationApproval) {
  const checks = targets.flatMap((target) =>
    findDataValidationsInRange(book, target.sheetId, target.address).map(
      (rule) => ({
        ...target,
        rule: rule,
      }),
    ),
  );
  if (!checks.length) return 0;
  if (checks.length > validationLimits.checks)
    throw Error(
      "Data validation exceeds the 100,000-check budget; split the batch.",
    );
  const evaluateRule = evaluateValidationRule(book),
    issues = checks
      .map((check) => evaluateRule(check.sheetId, check.address, check.rule))
      .filter((d) => !!d);
  if (!issues.length) return 0;
  const batchSignature = JSON.stringify({ book: book, issues: issues });
  if (
    validationApproval &&
    r5.get(validationApproval) === batchSignature &&
    issues.every((issue) => issue.style !== "stop" && !issue.evaluationError)
  )
    return (r5.delete(validationApproval), issues.length);
  throw validationApproval
    ? Error(
        "Validation approval is stale or invalid. Review the complete batch again.",
      )
    : new DataValidationError(book, issues);
}
function summarizeValidationInspection(workbook, sheetId) {
  const summary = {
      workbookId: workbook.id,
      revision: workbook.revision,
      checks: 0,
      invalidCells: 0,
      violations: 0,
      issues: [],
      complete: true,
      detailsTruncated: false,
      message: "",
    },
    evaluateRule = evaluateValidationRule(workbook),
    invalidCellKeys = new Set();
  let customCheckCount = 0;
  e: for (const worksheet of workbook.sheets.filter(
    (sheet) => !sheetId || sheet.id === sheetId,
  ))
    for (const validation of worksheet.validations ?? []) {
      const rangeAddress = parseRangeAddress(validation.range);
      for (
        let firstRow = rangeAddress.firstRow;
        firstRow <= rangeAddress.lastRow;
        firstRow++
      )
        for (
          let firstCol = rangeAddress.firstCol;
          firstCol <= rangeAddress.lastCol;
          firstCol++
        ) {
          if (
            summary.checks >= validationLimits.checks ||
            (validation.kind === "custom" &&
              ++customCheckCount > validationLimits.customChecks)
          ) {
            summary.complete = false;
            break e;
          }
          const cellAddress = `${columnIndexToLetters(firstCol)}${firstRow + 1}`;
          summary.checks++;
          const issue = evaluateRule(worksheet.id, cellAddress, validation);
          issue &&
            (invalidCellKeys.add(`${worksheet.id}:${cellAddress}`),
            summary.violations++,
            summary.issues.length < validationLimits.displayedIssues &&
              summary.issues.push(issue));
        }
    }
  return (
    (summary.invalidCells = invalidCellKeys.size),
    (summary.detailsTruncated = summary.violations > summary.issues.length),
    (summary.message = `${summary.complete ? "" : "Incomplete inspection: at least "}${summary.invalidCells} invalid cells; ${summary.violations} rule violations across ${summary.checks} checks.${summary.complete ? "" : " The 100,000-check or 5,000-custom-formula limit was reached. Narrow validation ranges and inspect again."}${summary.detailsTruncated ? " Showing the first 500 violations." : ""}`),
    FB.set(summary, {
      source: JSON.stringify(workbook),
      content: JSON.stringify(summary),
    }),
    summary
  );
}
function cU(workbook, report, values, calculationResult) {
  const spillsByKey = new Map(
      calculationResult.spills.map((spill) => [
        `${spill.sheetId}:${spill.address}`,
        spill,
      ]),
    ),
    spillMetadataByKey = new Map(
      calculationResult.spillMetadata.map((metadataEntry) => [
        `${metadataEntry.sheetId}:${metadataEntry.address}`,
        metadataEntry,
      ]),
    ),
    cellKey = (sheetId, cellAddress) =>
      `${sheetId}:${cellAddress.replace(/\$/g, "").toUpperCase()}`,
    sheetsById = new Map(workbook.sheets.map((sheet) => [sheet.id, sheet])),
    getTypedValue = (sheetId, cellAddress) => {
      const valueKey = cellKey(sheetId, cellAddress),
        storedValue = values.get(valueKey) ?? spillsByKey.get(valueKey)?.result;
      if (storedValue) return { ...storedValue };
      const sheet = sheetsById.get(sheetId);
      if (!sheet) return { type: "error", value: "#REF!" };
      try {
        const parsedAddress = _t(cellAddress);
        if (
          parsedAddress.row >= sheet.rowCount ||
          parsedAddress.col >= sheet.colCount
        )
          return { type: "error", value: "#REF!" };
      } catch {
        return { type: "error", value: "#REF!" };
      }
      const rawValue =
        sheet.cells[cellAddress.replace(/\$/g, "").toUpperCase()]?.raw ?? null;
      return {
        type:
          rawValue === null
            ? "blank"
            : typeof rawValue == "string"
              ? "text"
              : typeof rawValue == "boolean"
                ? "boolean"
                : "number",
        value: rawValue,
      };
    },
    getValue = (sheetId, cellAddress) =>
      getTypedValue(sheetId, cellAddress).value;
  return (
    Object.assign(getValue, {
      getTypedValue: getTypedValue,
      getDateContext: () => ({ ...report.dateContext }),
      getResourceReport: () => ({ ...report.resources }),
      getOperationalErrors: () => [...calculationResult.operationalErrors],
      getReferenceDiagnostics: () =>
        structuredClone(calculationResult.references),
      getNamedFormulaDiagnostics: () =>
        structuredClone(calculationResult.names),
      getFinancialDiagnostics: () =>
        structuredClone(calculationResult.financial),
      getSpillEntries: (sheetId) =>
        calculationResult.spills
          .filter((spill) => spill.sheetId === sheetId)
          .map(({ address, anchor }) => ({ address, anchor })),
      getSpillAnchor: (sheetId, cellAddress) =>
        spillsByKey.get(cellKey(sheetId, cellAddress))?.anchor,
      getSpillRange: (sheetId, cellAddress) =>
        spillMetadataByKey.get(cellKey(sheetId, cellAddress))?.range,
      getSpillDiagnostic: (sheetId, cellAddress) =>
        structuredClone(
          spillMetadataByKey.get(cellKey(sheetId, cellAddress))?.diagnostic,
        ),
      peekSpill: (sheetId, cellAddress) => {
        const spillInfo = spillMetadataByKey.get(cellKey(sheetId, cellAddress));
        return spillInfo
          ? {
              range: spillInfo.range,
              diagnostic: structuredClone(spillInfo.diagnostic),
            }
          : {};
      },
      evaluateFormula: (sheetId, cellAddress, formulaText) =>
        evaluateWorkbookFormulas(workbook).evaluateFormula(
          sheetId,
          cellAddress,
          formulaText,
        ),
    }),
    getValue
  );
}
function getIsolatedCalculation(workbook) {
  const existingRecord = calculatorRecords.get(workbook);
  if (
    existingRecord?.report.execution?.state === "pending" &&
    isCalculationRecordCurrent(existingRecord, workbook)
  )
    return {
      calculator: existingRecord.calculator,
      report: existingRecord.report,
    };
  const monotonicClockMs = readMonotonicClockMs(),
    cachedCalculationSnapshot = getCachedCalculationSnapshot(workbook),
    previousRecord =
      calculatorRecords.get(workbook) ?? carriedCalculatorRecords.get(workbook),
    values = new Map(),
    dependencies = new Map(),
    reportCellValueMap = buildReportCellValueMap(previousRecord),
    lastKnownValues = new Map();
  for (const formulaKey of cachedCalculationSnapshot.formulas.keys()) {
    values.set(formulaKey, pendingFormulaValue);
    const lastKnownValue = reportCellValueMap.get(formulaKey);
    lastKnownValue && lastKnownValues.set(formulaKey, lastKnownValue);
  }
  let cellReports;
  const baseReport = {
    workbookId: workbook.id,
    revision: workbook.revision,
    run: (previousRecord?.report.run ?? 0) + 1,
    settings: cachedCalculationSnapshot.settings,
    resources: createCalculationBudgetTracker(workbook).report(),
    volatile: {
      functions: cachedCalculationSnapshot.volatile,
      policy: calculationFeatureSupport,
    },
    dateContext: resolveLocalDateTime(workbook.datePolicy),
    reason: "Waiting for isolated calculation",
    formulaCount: cachedCalculationSnapshot.formulas.size,
    evaluated: 0,
    reused: 0,
    pending: cachedCalculationSnapshot.formulas.size,
    elapsedMs: readMonotonicClockMs() - monotonicClockMs,
    dependencyEdges: 0,
    cycles: [],
    iteration: {
      status: cachedCalculationSnapshot.settings.iterative
        ? "not-needed"
        : "off",
      passes: 0,
      maxDelta: null,
      message: "Waiting for isolated calculation",
    },
    execution: {
      mode: "worker",
      state: "pending",
      message:
        "Calculation is running separately; source inputs remain editable.",
    },
    limits: {
      dependencyEdges: j5,
      cooperativeTimeBudgetMs: cachedCalculationSnapshot.settings.timeBudgetMs,
      hardInterrupt: true,
    },
  };
  deepFreeze(baseReport);
  const report = Object.freeze({
    ...baseReport,
    get cells() {
      return (
        cellReports ||
          ((cellReports = [...cachedCalculationSnapshot.formulas].map(
            ([cellKey, formulaEntry]) => ({
              ...formulaEntry,
              result: { ...pendingFormulaValue },
              status: "pending",
              evaluations: 0,
              dependencies: [],
              ...(lastKnownValues.has(cellKey)
                ? { lastKnown: lastKnownValues.get(cellKey) }
                : {}),
            }),
          )),
          deepFreeze(cellReports)),
        cellReports
      );
    },
  });
  reportCellValueMaps.set(report, lastKnownValues);
  const spills = [];
  if (previousRecord)
    for (const sheet of workbook.sheets)
      for (const spillEntry of previousRecord.calculator.getSpillEntries(
        sheet.id,
      ))
        cachedCalculationSnapshot.formulas.has(
          `${sheet.id}:${spillEntry.anchor}`,
        ) &&
          spills.push({
            ...spillEntry,
            sheetId: sheet.id,
            result: { ...pendingFormulaValue },
          });
  const calculator = cU(workbook, report, values, {
    spills: spills,
    spillMetadata: [],
    references: [],
    names: [],
    financial: [],
    operationalErrors: cachedCalculationSnapshot.formulas.size
      ? ["#PENDING!"]
      : [],
  });
  return (
    calculatorRecords.set(workbook, {
      ...cachedCalculationSnapshot,
      calculator: calculator,
      report: report,
      values: values,
      dependencies: dependencies,
    }),
    carriedCalculatorRecords.delete(workbook),
    { calculator: calculator, report: report }
  );
}
function acceptWorkerCalculationResult(workbook, report, workerResult) {
  const existingRecord = calculatorRecords.get(workbook),
    cachedCalculationSnapshot = getCachedCalculationSnapshot(workbook);
  if (
    !existingRecord ||
    existingRecord.report !== report ||
    !isCalculationRecordCurrent(
      existingRecord,
      workbook,
      cachedCalculationSnapshot,
    )
  )
    return;
  if (
    workerResult.inputSignature !== getCalculationInputSignature(workbook) ||
    workerResult.report.workbookId !== workbook.id ||
    workerResult.report.revision !== workbook.revision ||
    workerResult.report.cells.length !== cachedCalculationSnapshot.formulas.size
  )
    throw Error("The worker result does not match the requested workbook.");
  const values = new Map(),
    dependencies = new Map();
  for (const cell of workerResult.report.cells) {
    const cellKey = `${cell.sheetId}:${cell.address}`;
    if (
      values.has(cellKey) ||
      cachedCalculationSnapshot.formulas.get(cellKey)?.formula !== cell.formula
    )
      throw Error(
        "The worker result contains stale or duplicate formula inputs.",
      );
    (values.set(cellKey, cell.result),
      dependencies.set(cellKey, new Set(cell.dependencies)));
  }
  const runReport = { ...workerResult.report, run: report.run };
  deepFreeze(runReport);
  const calculator = cU(workbook, runReport, values, workerResult);
  return (
    calculatorRecords.set(workbook, {
      ...cachedCalculationSnapshot,
      calculator: calculator,
      report: runReport,
      values: values,
      dependencies: dependencies,
    }),
    { calculator: calculator, report: runReport }
  );
}
function runIterativeCalculation(
  workbook,
  report,
  errorCode,
  errorMessage,
  wallElapsedMs,
) {
  if (errorCode === "#PENDING!") {
    const existingRecord = calculatorRecords.get(workbook);
    if (
      !existingRecord ||
      existingRecord.report !== report ||
      !isCalculationRecordCurrent(existingRecord, workbook)
    )
      return;
    const cancelledExecution = Object.freeze({
        mode: "worker",
        state: "cancelled",
        message: errorMessage,
        wallElapsedMs: wallElapsedMs,
      }),
      cancelledReport = Object.freeze(
        Object.defineProperties(
          {},
          {
            ...Object.getOwnPropertyDescriptors(report),
            reason: { value: errorMessage, enumerable: true },
            execution: { value: cancelledExecution, enumerable: true },
            iteration: {
              value: Object.freeze({
                ...report.iteration,
                status: report.settings.iterative ? "failed" : "off",
                message: errorMessage,
                passes: 0,
                maxDelta: null,
              }),
              enumerable: true,
            },
          },
        ),
      ),
      reportCellValues = reportCellValueMaps.get(report);
    return (
      reportCellValues &&
        reportCellValueMaps.set(cancelledReport, reportCellValues),
      calculatorRecords.set(workbook, {
        ...existingRecord,
        report: cancelledReport,
      }),
      { calculator: existingRecord.calculator, report: cancelledReport }
    );
  }
  const cells = report.cells.map((cell) => ({
    ...cell,
    result: makeFormulaError(errorCode),
    status: "error",
    evaluations: 0,
    dependencies: [],
  }));
  return acceptWorkerCalculationResult(workbook, report, {
    inputSignature: getCalculationInputSignature(workbook),
    report: {
      ...report,
      cells: cells,
      pending: 0,
      evaluated: 0,
      reused: 0,
      dependencyEdges: 0,
      cycles: [],
      reason: errorMessage,
      execution: {
        mode: "worker",
        state: errorCode === "#TIMEOUT!" ? "timeout" : "failed",
        message: errorMessage,
        wallElapsedMs: wallElapsedMs,
      },
      limits: { ...report.limits, hardInterrupt: true },
      iteration: {
        ...report.iteration,
        status: report.settings.iterative ? "failed" : "off",
        message: errorMessage,
        passes: 0,
        maxDelta: null,
      },
    },
    spills: [],
    spillMetadata: [],
    references: [],
    names: [],
    financial: [],
    operationalErrors: [errorCode],
  });
}
function recalculateWorkbookFormulas(
  workbook,
  cachedCalculationSnapshot,
  previousCalculation,
  isFullRecalculation,
) {
  const monotonicClockMs = readMonotonicClockMs(),
    { formulas, settings } = cachedCalculationSnapshot,
    localDateTime = resolveLocalDateTime(workbook.datePolicy),
    calculationBudgetTracker = createCalculationBudgetTracker(workbook);
  calculationBudgetTracker.begin();
  const dirtyFormulaKeys = new Set(),
    dependencies = new Map(),
    values = new Map(),
    evaluationCounts = new Map();
  let reason = isFullRecalculation
    ? "Explicit full recalculation"
    : previousCalculation
      ? "Changed inputs and transitive dependents"
      : "First calculation of this snapshot";
  if (
    isFullRecalculation ||
    !previousCalculation ||
    cachedCalculationSnapshot.metadata !== previousCalculation.metadata ||
    ik(workbook, formulas) ||
    ik(workbook, previousCalculation.formulas)
  )
    (formulas.forEach((formula, formulaKey) =>
      dirtyFormulaKeys.add(formulaKey),
    ),
      previousCalculation &&
        !isFullRecalculation &&
        (reason =
          cachedCalculationSnapshot.metadata !== previousCalculation.metadata
            ? "Calculation settings or structural metadata changed"
            : "Conservative recalculation for names, dynamic references or arrays"));
  else {
    const affectedKeys = new Set(),
      previousInputs = cachedCalculationSnapshot.inputs(),
      currentInputs = previousCalculation.inputs();
    for (const inputKey of new Set([
      ...previousInputs.keys(),
      ...currentInputs.keys(),
    ]))
      previousInputs.get(inputKey) !== currentInputs.get(inputKey) &&
        affectedKeys.add(inputKey);
    (formulas.forEach((formula, formulaKey) => {
      (_O.test(formula.formula) ||
        !previousCalculation.values.has(formulaKey) ||
        ["#TIMEOUT!", "#LIMIT!"].includes(
          String(previousCalculation.values.get(formulaKey)?.value),
        )) &&
        affectedKeys.add(formulaKey);
    }),
      previousCalculation.report.cycles
        .flat()
        .forEach((cycleCell) => affectedKeys.add(cycleCell)));
    const dependentsByPrecedent = new Map();
    previousCalculation.dependencies.forEach((dependency, dependentKey) =>
      dependency.forEach((precedentKey) => {
        const dependentKeys =
          dependentsByPrecedent.get(precedentKey) ?? new Set();
        (dependentKeys.add(dependentKey),
          dependentsByPrecedent.set(precedentKey, dependentKeys));
      }),
    );
    const changeQueue = [...affectedKeys];
    for (let queueIndex = 0; queueIndex < changeQueue.length; queueIndex++)
      for (const dependentKey of dependentsByPrecedent.get(
        changeQueue[queueIndex],
      ) ?? [])
        affectedKeys.has(dependentKey) ||
          (affectedKeys.add(dependentKey), changeQueue.push(dependentKey));
    affectedKeys.forEach((affectedKey) => {
      formulas.has(affectedKey) && dirtyFormulaKeys.add(affectedKey);
    });
  }
  formulas.forEach((formula, formulaKey) => {
    !dirtyFormulaKeys.has(formulaKey) &&
      previousCalculation?.values.has(formulaKey) &&
      (values.set(formulaKey, previousCalculation.values.get(formulaKey)),
      dependencies.set(
        formulaKey,
        new Set(previousCalculation.dependencies.get(formulaKey)),
      ));
  });
  const baseValues = new Map(values),
    cycles = [],
    cycleCells = new Set();
  let isRunning = true,
    stopError,
    dependencyEdgeCount = [...dependencies.values()].reduce(
      (total, dependencySet) => total + dependencySet.size,
      0,
    );
  const check = () => (
      isRunning && !stopError && (stopError = calculationBudgetTracker.check()),
      isRunning &&
        !stopError &&
        readMonotonicClockMs() - monotonicClockMs > settings.timeBudgetMs &&
        (stopError = "#TIMEOUT!"),
      isRunning ? stopError : undefined
    ),
    evaluatePass = (sourceValues, cycleOverrides) => {
      const evaluationStack = [];
      return evaluateWorkbookFormulas(workbook, {}, undefined, {
        budget: calculationBudgetTracker,
        running: () => isRunning,
        dateContext: localDateTime,
        check: check,
        read(referencedKey) {
          const currentFormulaKey = evaluationStack.at(-1);
          if (!isRunning || !currentFormulaKey) return;
          const currentDependencies = dependencies.get(currentFormulaKey);
          currentDependencies.has(referencedKey) ||
            (dependencyEdgeCount >= j5
              ? (stopError = "#LIMIT!")
              : (currentDependencies.add(referencedKey),
                dependencyEdgeCount++));
        },
        value(cellKey, shouldUseOverrides) {
          return shouldUseOverrides && cycleOverrides?.has(cellKey)
            ? cycleOverrides.get(cellKey)
            : sourceValues.get(cellKey);
        },
        enter(enteredKey) {
          (evaluationStack.push(enteredKey),
            (dependencyEdgeCount -= dependencies.get(enteredKey)?.size ?? 0),
            dependencies.set(enteredKey, new Set()),
            evaluationCounts.set(
              enteredKey,
              (evaluationCounts.get(enteredKey) ?? 0) + 1,
            ));
        },
        leave(leftKey, result) {
          (values.set(leftKey, result), evaluationStack.pop());
        },
        cycle(cyclePath) {
          const closingKey = cyclePath.at(-1),
            cycleMembers = [
              ...new Set(cyclePath.slice(cyclePath.indexOf(closingKey), -1)),
            ];
          (cycleMembers.forEach((cycleMember) => cycleCells.add(cycleMember)),
            cycleMembers.length &&
              !cycles.some(
                (cycle) =>
                  cycle.length === cycleMembers.length &&
                  cycle.every((memberKey) => cycleMembers.includes(memberKey)),
              ) &&
              cycles.push(cycleMembers));
        },
      });
    },
    isManualPending = settings.mode === "manual" && !isFullRecalculation;
  isManualPending &&
    dirtyFormulaKeys.forEach((dirtyKey) => {
      (values.set(dirtyKey, pendingFormulaValue),
        baseValues.set(dirtyKey, pendingFormulaValue),
        dependencies.set(
          dirtyKey,
          new Set(previousCalculation?.dependencies.get(dirtyKey)),
        ));
    });
  let baseCalculator = evaluatePass(baseValues);
  const copyEvaluatedValues = () => {
    for (const [formulaKey, formula] of formulas)
      values.set(
        formulaKey,
        baseValues.get(formulaKey) ??
          baseCalculator.getTypedValue(formula.sheetId, formula.address),
      );
  };
  isManualPending || copyEvaluatedValues();
  const iteration = {
    status: settings.iterative ? "not-needed" : "off",
    passes: 0,
    maxDelta: null,
    message: settings.iterative
      ? "No circular calculation was required."
      : "Iteration is disabled. Circular references return #CYCLE!.",
  };
  if (!isManualPending && cycleCells.size && settings.iterative) {
    const iterationKeys = new Set(cycleCells),
      dependentsByKey = new Map();
    dependencies.forEach((dependency, dependentKey) =>
      dependency.forEach((precedentKey) => {
        const dependentSet = dependentsByKey.get(precedentKey) ?? new Set();
        (dependentSet.add(dependentKey),
          dependentsByKey.set(precedentKey, dependentSet));
      }),
    );
    const expansionQueue = [...cycleCells];
    for (
      let expansionIndex = 0;
      expansionIndex < expansionQueue.length;
      expansionIndex++
    )
      for (const dependentKey of dependentsByKey.get(
        expansionQueue[expansionIndex],
      ) ?? [])
        iterationKeys.has(dependentKey) ||
          (iterationKeys.add(dependentKey), expansionQueue.push(dependentKey));
    if (
      (values.forEach((value, cellKey) => {
        iterationKeys.has(cellKey) || baseValues.set(cellKey, value);
      }),
      (iteration.status = "failed"),
      ik(workbook, formulas) ||
        [...formulas.values()].some((formula) => _O.test(formula.formula)) ||
        cycleCells.size > 1e3)
    )
      iteration.message =
        "Iteration supports at most 1,000 deterministic scalar cycle cells. Names, arrays, dynamic references and volatile formulas must be removed from this calculation.";
    else {
      let cycleGuesses = new Map(
        [...cycleCells].map((cycleKey) => [
          cycleKey,
          { type: "number", value: 0 },
        ]),
      );
      iteration.message = `No convergence within ${settings.maxIterations} simultaneous passes from zero.`;
      for (
        let passIndex = 0;
        passIndex < settings.maxIterations && !check();
        passIndex++
      ) {
        baseCalculator = evaluatePass(baseValues, cycleGuesses);
        const nextGuesses = new Map();
        let maxDelta = 0,
          isAllNumeric = true;
        for (const cycleKey of cycleCells) {
          const cycleFormula = formulas.get(cycleKey),
            cycleValue = baseCalculator.getTypedValue(
              cycleFormula.sheetId,
              cycleFormula.address,
            );
          if (
            (nextGuesses.set(cycleKey, cycleValue),
            cycleValue.type !== "number" || !Number.isFinite(cycleValue.value))
          ) {
            isAllNumeric = false;
            break;
          }
          maxDelta = Math.max(
            maxDelta,
            Math.abs(
              Number(cycleValue.value) -
                Number(cycleGuesses.get(cycleKey).value),
            ),
          );
        }
        if (
          ((iteration.passes = passIndex + 1),
          (iteration.maxDelta = Number.isFinite(maxDelta) ? maxDelta : null),
          !isAllNumeric)
        ) {
          iteration.message =
            "The circular calculation produced a nonnumeric value or calculation error.";
          break;
        }
        if (((cycleGuesses = nextGuesses), maxDelta <= settings.maxChange)) {
          ((iteration.status = "converged"),
            (iteration.message = `Converged from zero in ${passIndex + 1} simultaneous passes; maximum change ${maxDelta}.`));
          break;
        }
      }
      iteration.status === "converged" &&
        !stopError &&
        cycleCells.forEach((cycleKey) =>
          baseValues.set(cycleKey, cycleGuesses.get(cycleKey)),
        );
    }
    ((iteration.status !== "converged" || stopError) &&
      cycleCells.forEach((cycleKey) =>
        baseValues.set(cycleKey, makeFormulaError(stopError ?? "#ITER!")),
      ),
      (baseCalculator = evaluatePass(baseValues)),
      copyEvaluatedValues());
  }
  (check(),
    stopError &&
      (dirtyFormulaKeys.forEach((dirtyKey) =>
        values.set(dirtyKey, makeFormulaError(stopError)),
      ),
      iteration.status === "converged" && (iteration.status = "failed"),
      cycleCells.size &&
        settings.iterative &&
        (iteration.message = `Calculation stopped with ${stopError}; no intermediate result is current.`)),
    (isRunning = false),
    calculationBudgetTracker.end());
  const calculator = (targetSheetId, targetAddress) =>
    calculator.getTypedValue(targetSheetId, targetAddress).value;
  Object.assign(calculator, baseCalculator);
  const toCellKey = (sheetId, address) =>
      `${sheetId}:${address.replace(/\$/g, "").toUpperCase()}`,
    pendingSpillAnchors = new Map();
  if (isManualPending && previousCalculation)
    for (const sheet of workbook.sheets)
      for (const spillEntry of previousCalculation.calculator.getSpillEntries(
        sheet.id,
      ))
        dirtyFormulaKeys.has(`${sheet.id}:${spillEntry.anchor}`) &&
          pendingSpillAnchors.set(
            `${sheet.id}:${spillEntry.address}`,
            spillEntry.anchor,
          );
  calculator.getTypedValue = (sheetId, address) => {
    const valueKey = toCellKey(sheetId, address);
    if (values.has(valueKey)) return { ...values.get(valueKey) };
    const cell = workbook.sheets.find((sheet) => sheet.id === sheetId)?.cells[
      address.replace(/\$/g, "").toUpperCase()
    ];
    if (
      cell &&
      cell.raw !== null &&
      (cell.literal || typeof cell.raw != "string" || !cell.raw.startsWith("="))
    )
      return {
        type:
          typeof cell.raw == "string"
            ? "text"
            : typeof cell.raw == "boolean"
              ? "boolean"
              : "number",
        value: cell.raw,
      };
    if (pendingSpillAnchors.has(valueKey)) return { ...pendingFormulaValue };
    const spillAnchor = baseCalculator.getSpillAnchor(sheetId, address);
    return spillAnchor &&
      values.get(`${sheetId}:${spillAnchor}`)?.type === "error"
      ? { ...values.get(`${sheetId}:${spillAnchor}`) }
      : baseCalculator.getTypedValue(sheetId, address);
  };
  const getSpillEntries = baseCalculator.getSpillEntries;
  ((calculator.getSpillEntries = (sheetId) => [
    ...getSpillEntries(sheetId),
    ...[...pendingSpillAnchors]
      .filter(([pendingKey]) => pendingKey.startsWith(`${sheetId}:`))
      .map(([pendingKey, anchor]) => ({
        address: pendingKey.slice(sheetId.length + 1),
        anchor: anchor,
      })),
  ]),
    (calculator.getOperationalErrors = () => [
      ...new Set([
        ...baseCalculator.getOperationalErrors(),
        ...[...values.values()]
          .filter(
            (cellValue) =>
              cellValue.type === "error" &&
              [
                "#PENDING!",
                "#TIMEOUT!",
                "#ITER!",
                "#LIMIT!",
                "#BLOCKED!",
                "#CYCLE!",
                "#ERROR!",
              ].includes(String(cellValue.value)),
          )
          .map((cellValue) => cellValue.value),
      ]),
    ]));
  const mergeDiagnostics = (previousDiagnostics, evaluatorDiagnostics) => [
      ...previousDiagnostics.filter(
        (diagnostic) =>
          !dirtyFormulaKeys.has(
            `${diagnostic.sheetId}:${diagnostic.address}`,
          ) && formulas.has(`${diagnostic.sheetId}:${diagnostic.address}`),
      ),
      ...evaluatorDiagnostics,
    ],
    previousReferenceDiagnostics =
      previousCalculation?.calculator.getReferenceDiagnostics() ?? [],
    previousNamedFormulaDiagnostics =
      previousCalculation?.calculator.getNamedFormulaDiagnostics() ?? [],
    previousFinancialDiagnostics =
      previousCalculation?.calculator.getFinancialDiagnostics() ?? [];
  ((calculator.getFinancialDiagnostics = () =>
    structuredClone(
      mergeDiagnostics(
        previousFinancialDiagnostics,
        baseCalculator.getFinancialDiagnostics(),
      ),
    )),
    (calculator.getReferenceDiagnostics = () =>
      mergeDiagnostics(
        previousReferenceDiagnostics,
        baseCalculator.getReferenceDiagnostics(),
      )),
    (calculator.getNamedFormulaDiagnostics = () =>
      mergeDiagnostics(
        previousNamedFormulaDiagnostics,
        baseCalculator.getNamedFormulaDiagnostics(),
      )));
  const previousCellReports = new Map(
      previousCalculation?.report.cells.map((cell) => [
        `${cell.sheetId}:${cell.address}`,
        cell,
      ]),
    ),
    financialByCell = new Map();
  for (const {
    sheetId,
    address,
    ...diagnosticDetails
  } of calculator.getFinancialDiagnostics()) {
    const cellKey = `${sheetId}:${address}`;
    financialByCell.set(cellKey, [
      ...(financialByCell.get(cellKey) ?? []),
      diagnosticDetails,
    ]);
  }
  const cells = [...formulas].map(([formulaKey, formula]) => {
    const cellResult = values.get(formulaKey) ?? pendingFormulaValue,
      previousResult = previousCalculation?.values.get(formulaKey);
    return {
      ...formula,
      ...(cellResult.value !== "#PENDING!" && financialByCell.has(formulaKey)
        ? { financial: financialByCell.get(formulaKey) }
        : {}),
      result: { ...cellResult },
      status:
        cellResult.type === "error"
          ? cellResult.value === "#PENDING!"
            ? "pending"
            : "error"
          : "current",
      ...(cellResult.value === "#PENDING!" &&
      previousResult &&
      previousResult.value !== "#PENDING!"
        ? {
            lastKnown: {
              revision: previousCalculation.report.revision,
              result: { ...previousResult },
            },
          }
        : {}),
      ...(cellResult.value === "#PENDING!" &&
      previousResult?.value === "#PENDING!" &&
      previousCellReports.get(formulaKey)?.lastKnown
        ? { lastKnown: previousCellReports.get(formulaKey).lastKnown }
        : {}),
      dependencies: [...(dependencies.get(formulaKey) ?? [])].sort(),
      evaluations: evaluationCounts.get(formulaKey) ?? 0,
    };
  });
  return {
    ...cachedCalculationSnapshot,
    calculator: calculator,
    values: values,
    dependencies: dependencies,
    report: {
      resources: calculationBudgetTracker.report(),
      volatile: {
        functions: [
          ...new Set(
            [...formulas.values()]
              .flatMap((formula) => findVolatileFunctionNames(formula.formula))
              .concat(
                (workbook.definedNames ?? []).flatMap((definedName) =>
                  findVolatileFunctionNames(definedName.refersTo),
                ),
              ),
          ),
        ].sort(),
        policy: calculationFeatureSupport,
      },
      execution: {
        mode: "cooperative",
        state:
          stopError === "#TIMEOUT!"
            ? "timeout"
            : stopError
              ? "failed"
              : isManualPending
                ? "pending"
                : "complete",
        message: stopError
          ? (calculationBudgetTracker.report().reason ??
            `Calculation stopped with ${stopError}.`)
          : "Synchronous calculation uses cooperative limits.",
      },
      dateContext: localDateTime,
      workbookId: workbook.id,
      revision: workbook.revision,
      run: (previousCalculation?.report.run ?? 0) + 1,
      settings,
      reason: reason,
      formulaCount: formulas.size,
      evaluated: cells.filter((cell) => cell.evaluations > 0).length,
      reused: formulas.size - dirtyFormulaKeys.size,
      pending: cells.filter((cell) => cell.status === "pending").length,
      elapsedMs: Math.max(0, readMonotonicClockMs() - monotonicClockMs),
      dependencyEdges: [...dependencies.values()].reduce(
        (total, dependencySet) => total + dependencySet.size,
        0,
      ),
      cycles: cycles,
      iteration: iteration,
      cells: cells,
      limits: {
        dependencyEdges: j5,
        cooperativeTimeBudgetMs: settings.timeBudgetMs,
        hardInterrupt: false,
      },
    },
  };
}
function getWorkbookCalculation(workbook) {
  const cachedCalculationSnapshot = getCachedCalculationSnapshot(workbook),
    calculatorRecord = calculatorRecords.get(workbook);
  if (
    isCalculationRecordCurrent(
      calculatorRecord,
      workbook,
      cachedCalculationSnapshot,
    )
  )
    return {
      calculator: calculatorRecord.calculator,
      report: calculatorRecord.report,
    };
  const recalculated = recalculateWorkbookFormulas(
    workbook,
    cachedCalculationSnapshot,
    calculatorRecord ? undefined : carriedCalculatorRecords.get(workbook),
    false,
  );
  return (
    deepFreeze(recalculated.report),
    calculatorRecords.set(workbook, recalculated),
    carriedCalculatorRecords.delete(workbook),
    { calculator: recalculated.calculator, report: recalculated.report }
  );
}
function calculateWorkbookWithSettings(workbook, settings = {}, t) {
  return !t &&
    !Object.keys(settings).length &&
    (workbook.calculation ||
      workbook.datePolicy ||
      hasCalculatorRecord(workbook) ||
      workbookUsesNowOrToday(workbook) ||
      workbookHasVolatileFunctions(workbook))
    ? getWorkbookCalculation(workbook).calculator
    : evaluateWorkbookFormulas(workbook, settings, t);
}
function updateConditionalFormatRules(workbook, sheetId, rules, options = {}) {
  if (options.readOnly) throw Error("This workbook is read-only.");
  const sheet = workbook.sheets.find((sheet) => sheet.id === sheetId);
  if (!sheet) throw Error("Worksheet no longer exists.");
  if (sheet.protection?.enabled)
    throw Error("Unprotect the worksheet before changing conditional formats.");
  const conditionalRules = normalizeConditionalRules(rules),
    existingRules = sheet.conditionalFormats ?? [],
    changes = [];
  for (const ruleId of new Set([
    ...existingRules.map((existingRule) => existingRule.id),
    ...conditionalRules.map((conditionalRule) => conditionalRule.id),
  ])) {
    const beforeIndex = existingRules.findIndex(
        (storedRule) => storedRule.id === ruleId,
      ),
      afterIndex = conditionalRules.findIndex(
        (conditionalRule) => conditionalRule.id === ruleId,
      ),
      previousRule = existingRules[beforeIndex],
      updatedRule = conditionalRules[afterIndex];
    (JSON.stringify(previousRule) !== JSON.stringify(updatedRule) ||
      beforeIndex !== afterIndex) &&
      changes.push({
        id: ruleId,
        ...(previousRule
          ? { before: structuredClone(previousRule), beforeIndex: beforeIndex }
          : {}),
        ...(updatedRule
          ? { after: structuredClone(updatedRule), afterIndex: afterIndex }
          : {}),
      });
  }
  const workbookWithSettings = calculateWorkbookWithSettings(workbook),
    beforeEvaluator = evaluateConditionalFormatRules(
      sheet,
      workbookWithSettings,
    ),
    afterEvaluator = evaluateConditionalFormatRules(
      { ...sheet, conditionalFormats: conditionalRules },
      workbookWithSettings,
    ),
    previewAddresses = new Set();
  for (const rule of [...conditionalRules, ...existingRules]) {
    const rangeAddress = parseRangeAddress(rule.range);
    if (
      (previewAddresses.add(
        `${columnIndexToLetters(rangeAddress.firstCol)}${rangeAddress.firstRow + 1}`,
      ),
      previewAddresses.size >= 50)
    )
      break;
  }
  if (previewAddresses.size < 50) {
    for (const cell of Object.keys(sheet.cells))
      if (
        ([...conditionalRules, ...existingRules].some((formatRule) =>
          isCellInsideRange(parseRangeAddress(formatRule.range), _t(cell)),
        ) && previewAddresses.add(cell),
        previewAddresses.size >= 50)
      )
        break;
  }
  const samples = [...previewAddresses].map((address) => ({
    address: address,
    value: workbookWithSettings(sheetId, address),
    before: beforeEvaluator(address),
    after: afterEvaluator(address),
  }));
  return {
    workbookId: workbook.id,
    revision: workbook.revision,
    snapshotGuard: JSON.stringify(workbook),
    sheetId: sheetId,
    rules: conditionalRules,
    changes: changes,
    samples: samples,
    errors: afterEvaluator.errors(),
    previewLimit: 50,
  };
}
function refreshCalculationReport(workbook) {
  const cachedCalculationSnapshot = getCachedCalculationSnapshot(workbook),
    currentRecord = calculatorRecords.get(workbook),
    recalculatedRecord = recalculateWorkbookFormulas(
      workbook,
      cachedCalculationSnapshot,
      isCalculationRecordCurrent(
        currentRecord,
        workbook,
        cachedCalculationSnapshot,
      )
        ? currentRecord
        : undefined,
      true,
    );
  return (
    deepFreeze(recalculatedRecord.report),
    calculatorRecords.set(workbook, recalculatedRecord),
    carriedCalculatorRecords.delete(workbook),
    {
      calculator: recalculatedRecord.calculator,
      report: recalculatedRecord.report,
    }
  );
}
function buildCalculationEvidence(workbook, report) {
  const workbookCalculation = getWorkbookCalculation(workbook);
  if (workbookCalculation.report !== report)
    throw Error(
      "The workbook or calculation run changed. Review current calculation results before exporting evidence.",
    );
  const importedCalculationCaches = workbook.importedCalculationCaches,
    entries = (importedCalculationCaches?.entries ?? []).map((cacheEntry) => {
      const sheet = workbook.sheets.find(
          (sheet) => sheet.id === cacheEntry.sheetId,
        ),
        cell = sheet?.cells[cacheEntry.address],
        currentTypedValue = sheet
          ? workbookCalculation.calculator.getTypedValue(
              cacheEntry.sheetId,
              cacheEntry.address,
            )
          : undefined,
        isCacheMatch =
          cacheEntry.cache &&
          currentTypedValue &&
          cacheEntry.cache.type === currentTypedValue.type &&
          (typeof cacheEntry.cache.value == "number" &&
          typeof currentTypedValue.value == "number"
            ? Math.abs(cacheEntry.cache.value - currentTypedValue.value) <=
              1e-10 *
                Math.max(
                  1,
                  Math.abs(cacheEntry.cache.value),
                  Math.abs(currentTypedValue.value),
                )
            : cacheEntry.cache.value === currentTypedValue.value),
        evidenceStatus = sheet
          ? cell?.literal
            ? "neutralized-or-literal"
            : cell?.raw !== cacheEntry.formula
              ? "formula-changed"
              : currentTypedValue?.type === "error" &&
                  currentTypedValue.value === "#PENDING!"
                ? "pending"
                : cacheEntry.cache
                  ? isCacheMatch
                    ? "matched"
                    : "different"
                  : "missing"
          : "source-sheet-removed";
      return {
        ...cacheEntry,
        ...(currentTypedValue ? { current: currentTypedValue } : {}),
        status: evidenceStatus,
        verifiedByExcel: false,
      };
    });
  return JSON.stringify(
    {
      schemaVersion: 1,
      report: report,
      imported: importedCalculationCaches
        ? { ...importedCalculationCaches, entries: entries }
        : null,
      boundaries: [
        "Imported caches are unverified source evidence, never current-value fallbacks.",
        "Matching caches do not prove Excel parity or source-file authenticity.",
        "Calculation timings and iteration counts are local runtime evidence.",
        report.limits.hardInterrupt
          ? "This calculation ran in a terminable worker. Timer delivery is subject to runtime scheduling."
          : "This synchronous calculation used cooperative time checks, not a hard worker interrupt.",
        "Working-data accounting is a bounded allocation allowance, not an exact heap/RSS measurement; runtime overhead is outside it.",
        "Diagnostic formula evaluation is separate from the committed revision calculation.",
      ],
    },
    null,
    2,
  );
}
function parseReportRange(range) {
  if (
    !/^\$?[A-Za-z]{1,3}\$?[1-9]\d{0,6}(?::\$?[A-Za-z]{1,3}\$?[1-9]\d{0,6})?$/.test(
      range,
    )
  )
    throw Error("Choose one explicit report range.");
  const addressParts = range.replaceAll("$", "").toUpperCase().split(":"),
    rangeAddress = parseRangeAddress(range);
  if (
    addressParts[0] !==
      `${columnIndexToLetters(rangeAddress.firstCol)}${rangeAddress.firstRow + 1}` ||
    (addressParts[1] &&
      addressParts[1] !==
        `${columnIndexToLetters(rangeAddress.lastCol)}${rangeAddress.lastRow + 1}`)
  )
    throw Error("Report ranges cannot be reversed.");
  if (
    rangeAddress.lastCol - rangeAddress.firstCol + 1 > 20 ||
    rangeAddress.lastRow - rangeAddress.firstRow + 1 > 500 ||
    (rangeAddress.lastCol - rangeAddress.firstCol + 1) *
      (rangeAddress.lastRow - rangeAddress.firstRow + 1) >
      2e3
  )
    throw Error(
      "Reports support at most 500 rows, 20 columns and 2,000 cells.",
    );
  return `${columnIndexToLetters(rangeAddress.firstCol)}${rangeAddress.firstRow + 1}:${columnIndexToLetters(rangeAddress.lastCol)}${rangeAddress.lastRow + 1}`;
}
const PO = object({
    column: number().int().min(0).max(16383),
    type: _enum(["value", "text", "number", "date"]),
    operator: _enum([
      "equals",
      "notEquals",
      "contains",
      "notContains",
      "startsWith",
      "endsWith",
      "wildcard",
      "gt",
      "gte",
      "lt",
      "lte",
      "blank",
      "notBlank",
    ]),
    value: union([
      string().max(1024),
      number().finite(),
      boolean(),
      _null(),
    ]).optional(),
  }).strict(),
  filterRequestSchema = object({
    range: createBoundedTextSchema(40),
    columns: array(
      object({
        column: number().int().min(0).max(16383),
        values: array(
          union([
            createBoundedTextSchema(32767),
            number().finite(),
            boolean(),
            _null(),
          ]),
        )
          .max(1e3)
          .optional(),
        conditions: array(PO).max(128).optional(),
        join: _enum(["and", "or"]).optional(),
      }).strict(),
    ).max(20),
    advanced: array(array(PO).max(128)).max(128).optional(),
    caseSensitive: boolean().optional(),
    unique: boolean().optional(),
  }).strict(),
  kU = object({
    name: string().trim().min(1).max(100),
    kind: _enum(["range", "object"]),
    sheetId: nonEmptyLabelSchema,
    range: createBoundedTextSchema(40).transform(parseReportRange),
    objectId: nonEmptyLabelSchema.optional(),
    mode: _enum(["snapshot", "linked"]),
    population: _enum(["full", "visible"]),
    filter: filterRequestSchema.optional(),
  }).strict(),
  cellSnapshotSchema = object({
    address: string().regex(/^[A-Z]{1,3}[1-9]\d{0,6}$/),
    value: union([
      createBoundedTextSchema(32767),
      number().finite(),
      boolean(),
      _null(),
    ]),
    text: createBoundedTextSchema(32767),
    style: unknown().optional(),
  }).strict(),
  calculationReportSchema = object({
    engine: literal(reportOutputFormatVersion),
    id: nonEmptyLabelSchema,
    request: kU,
    workbookId: nonEmptyLabelSchema,
    workbookName: createBoundedTextSchema(200),
    revision: number().int().nonnegative(),
    sheetName: createBoundedTextSchema(31),
    calculatedAt: AU,
    sourceHash: reportHashSchema,
    includeHidden: boolean(),
    authorizedSheetIds: array(nonEmptyLabelSchema).max(1e3),
    sourceCells: array(
      object({
        sheetId: nonEmptyLabelSchema,
        address: string().regex(/^[A-Z]{1,3}[1-9]\d{0,6}$/),
      }).strict(),
    ).max(4e3),
    rows: array(
      object({
        row: number().int().min(1).max(1048576),
        cells: array(cellSnapshotSchema).max(20),
      }).strict(),
    ).max(500),
    columns: array(number().int().min(0).max(16383)).max(20),
    widths: array(number().finite().min(40).max(600)).max(20),
    excludedRows: array(number().int().min(1).max(1048576)).max(500),
    object: unknown().optional(),
    warnings: array(createBoundedTextSchema(2e3)).max(50),
    hash: reportHashSchema,
  }).strict(),
  importedCacheEvidenceSchema = object({
    id: nonEmptyLabelSchema,
    fileName: createBoundedTextSchema(255),
    fileBytes: number()
      .int()
      .min(1)
      .max(10 * 1024 * 1024),
    sourceSha256: reportHashSchema,
    inspectedAt: AU,
    execution: literal("never"),
    items: array(
      object({
        path: createBoundedTextSchema(4096),
        kind: createBoundedTextSchema(100),
        bytes: number()
          .int()
          .nonnegative()
          .max(50 * 1024 * 1024),
        sha256: reportHashSchema.optional(),
        execution: literal("never"),
        note: createBoundedTextSchema(2e3),
      }).strict(),
    )
      .min(1)
      .max(1e3),
    warnings: array(createBoundedTextSchema(2e3)).max(100),
  }).strict();
function validateObjectReportRequest(request) {
  const parsedRequest = kU.parse(request);
  if ((parsedRequest.kind === "object") != !!parsedRequest.objectId)
    throw Error("Select an existing drawing for an object report.");
  if (
    parsedRequest.filter &&
    parseReportRange(parsedRequest.filter.range) !== parsedRequest.range
  )
    throw Error("The report range must match its filter range.");
  return parsedRequest;
}

export {
  validateObjectReportRequest,
  evaluateWorkbookFormulas,
  parseReportRange,
  getWorkbookCalculation,
  getIsolatedCalculation,
  updateConditionalFormatRules,
  summarizeValidationInspection,
  importedCacheEvidenceSchema,
  calculationReportSchema,
  calculateWorkbookWithSettings,
  buildCalculationEvidence,
  runIterativeCalculation,
  refreshCalculationReport,
  acceptWorkerCalculationResult,
  checkDataValidationBatch,
};

/* Also used by the workers. */
export {
  PO,
  cellSnapshotSchema,
  filterRequestSchema,
  kU,
  recalculateWorkbookFormulas,
};

/* Also used by the workers. */
export { evaluateValidationRule };
