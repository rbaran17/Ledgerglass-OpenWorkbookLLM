import { flattenNestedValues } from "../02-text-and-financial-functions.js";
import { isPlainObjectValue } from "./01-cell-styles-and-themes.js";
import { isNullOrEmptyValue } from "./03-defined-names.js";
import {
  evaluateDateFunction,
  notDateFunctionSentinel,
  resolveLocalDateTime,
  dateSystemSerialOffset,
  parseDateText,
  findVolatileFunctionNames,
  isDeepFrozenValue,
  getExtensionCacheEntries,
  normalizeDatePolicy,
  dateSystemMaxSerial,
  fingerprintCanonicalValue,
} from "../01-schemas-date-functions-and-validation.js";
import {
  fileImportLimits,
  evaluateFinancialFunction,
  notFinancialSentinel,
  evaluateTextFunction,
  notTextFunctionSentinel,
  broadcastOverArrays,
} from "../02-text-and-financial-functions.js";
import {
  string,
  number,
  boolean,
  array,
  strictObject,
  discriminatedUnion,
  literal,
  union,
  enum as _enum,
  null as _null,
} from "zod";
import {
  validateNumberFormatCode,
  numberFormatToIntlOptions,
} from "./01-cell-styles-and-themes.js";
import {
  notRegressionSentinel,
  notArithmeticSentinel,
  notSpecialistSentinel,
  notInferenceSentinel,
  evaluateArithmeticFunction,
  notStatisticalSentinel,
  sumDecimalPartsToNumber,
  exactAverageOfNumbers,
  evaluateSignificanceTestFunction,
  notOwnedLookupSentinel,
  evaluateSpecialistFunction,
  evaluateLookupFunction,
  evaluateRegressionFunction,
  evaluateLogicalFunction,
  notModernFunctionSentinel,
  evaluateStatisticalFunction,
  notOwnedLogicalSentinel,
  formatCellValueWithNumberFormat,
  matrixResultFunctionNames,
} from "./02-formula-function-evaluators.js";
import {
  definedNameScopeKey,
  evaluateDefinedNameFunction,
  validateDefinedName,
  replaceTextSpans,
  m5,
  pO,
} from "./03-defined-names.js";
import {
  findFormulaErrorArg,
  builtinFunctionNameSet,
  evaluateDatabaseFunction,
  findScopedDefinedName,
  formulaErrorFactory,
  createFormulaError,
  notDatabaseFunctionMarker,
  notReferenceFunctionMarker,
  resolveDefinedNameValue,
  isFormulaError,
} from "./04-pivot-tables-and-pictures.js";

function applyDefinedNameOperation(workbook, operation, options = {}) {
  if (options.readOnly) throw Error("This workbook is read-only.");
  if (workbook.structureProtection?.enabled)
    throw Error(
      "Workbook structure is protected. Unprotect it before changing names.",
    );
  const definedNames = validateDefinedName(workbook.definedNames, workbook),
    existingName =
      operation.kind === "create"
        ? undefined
        : definedNames.find(
            (candidate) =>
              candidate.name.toLowerCase() === operation.name.toLowerCase() &&
              candidate.scopeSheetId === operation.scopeSheetId,
          );
  if (operation.kind !== "create" && !existingName)
    throw Error("The defined name was not found. Reopen Name Manager.");
  let newDefinition;
  if (operation.kind === "create" || operation.kind === "update")
    newDefinition = validateDefinedName([operation.definition], workbook)[0];
  else if (operation.kind === "rename")
    newDefinition = validateDefinedName(
      [{ ...existingName, name: operation.newName }],
      workbook,
    )[0];
  else if (operation.kind !== "delete") throw Error("Unknown name operation.");
  if (
    operation.kind === "update" &&
    (existingName.name !== newDefinition.name ||
      existingName.scopeSheetId !== newDefinition.scopeSheetId)
  )
    throw Error(
      "Use Rename to change a name. Existing name scope cannot be changed; create a separate name instead.",
    );
  if (
    newDefinition &&
    workbook.sheets.some((sheet) =>
      sheet.tables?.some(
        (table) =>
          table.name.toLowerCase() === newDefinition.name.toLowerCase(),
      ),
    )
  )
    throw Error("A table already uses this name.");
  const next = structuredClone(workbook),
    existingScopeKey = existingName
      ? definedNameScopeKey(existingName)
      : undefined;
  ((next.definedNames = validateDefinedName(
    [
      ...definedNames.filter(
        (otherEntry) => definedNameScopeKey(otherEntry) !== existingScopeKey,
      ),
      ...(newDefinition ? [newDefinition] : []),
    ],
    workbook,
  )),
    existingName &&
      newDefinition &&
      (next.definedNames = definedNames.map((existingEntry) =>
        definedNameScopeKey(existingEntry) === existingScopeKey
          ? newDefinition
          : existingEntry,
      )));
  const affectedSheetId =
      newDefinition?.scopeSheetId ?? existingName?.scopeSheetId,
    assertSheetUnprotected = (sheetId) => {
      if (
        workbook.sheets.find((sheet) => sheet.id === sheetId)?.protection
          ?.enabled
      )
        throw Error(
          "An affected worksheet is protected. Unprotect it before changing names.",
        );
    };
  affectedSheetId && assertSheetUnprotected(affectedSheetId);
  const changes = [],
    dependencies = [];
  let scanBudgetUsed = 0;
  const p = true,
    scanReferences = (formulaText) => {
      const scannedReferences = m5(formulaText, p);
      if (
        ((scanBudgetUsed += scannedReferences.length + 1), scanBudgetUsed > 1e6)
      )
        throw Error(
          "The name dependency scan exceeds the supported budget. Split the workbook before changing names.",
        );
      if (
        scannedReferences.some(
          (scannedReference) =>
            scannedReference.qualifier &&
            !workbook.sheets.some(
              (sheet) =>
                sheet.name.toLowerCase() ===
                scannedReference.qualifier.toLowerCase(),
            ),
        )
      )
        throw Error(
          "An unknown or external worksheet qualifier cannot be safely migrated.",
        );
      return scannedReferences;
    },
    dependsOnTarget = (refersToText, scopeSheetId, dependencyChain = []) => {
      if (dependencyChain.length > 32)
        throw Error("The name dependency chain exceeds 32 levels.");
      for (const reference of scanReferences(refersToText)) {
        const scopedDefinedName = findScopedDefinedName(
            reference,
            workbook,
            scopeSheetId,
          ),
          nextDefinedName =
            operation.kind === "create"
              ? findScopedDefinedName(reference, next, scopeSheetId)
              : undefined;
        if (
          nextDefinedName &&
          definedNameScopeKey(nextDefinedName) ===
            definedNameScopeKey(newDefinition) &&
          (!scopedDefinedName ||
            definedNameScopeKey(scopedDefinedName) !==
              definedNameScopeKey(nextDefinedName))
        )
          return true;
        if (!scopedDefinedName) {
          if (
            operation.kind === "create" &&
            findScopedDefinedName(reference, next, scopeSheetId) &&
            definedNameScopeKey(
              findScopedDefinedName(reference, next, scopeSheetId),
            ) === definedNameScopeKey(newDefinition)
          )
            return true;
          continue;
        }
        if (definedNameScopeKey(scopedDefinedName) === existingScopeKey)
          return true;
        const dependencyScopeKey = definedNameScopeKey(scopedDefinedName);
        if (
          !dependencyChain.includes(dependencyScopeKey) &&
          dependsOnTarget(
            scopedDefinedName.refersTo,
            scopedDefinedName.scopeSheetId ??
              (reference.qualifier
                ? workbook.sheets.find(
                    (sheet) =>
                      sheet.name.toLowerCase() ===
                      reference.qualifier.toLowerCase(),
                  ).id
                : scopeSheetId),
            [...dependencyChain, dependencyScopeKey],
          )
        )
          return true;
      }
      return false;
    },
    rewriteFormula = (before, sheetIds, target) => {
      const beforeReferences = scanReferences(before);
      if (
        sheetIds.some((candidateSheetId) =>
          dependsOnTarget(before, candidateSheetId),
        )
      ) {
        if (
          (dependencies.push({
            label: target.label,
            ...(target.sheetId ? { sheetId: target.sheetId } : {}),
            ...(target.address ? { address: target.address } : {}),
          }),
          dependencies.length > pO)
        )
          throw Error(
            "A name operation may affect at most 500 dependent locations. Split the operation into smaller workbooks.",
          );
        if (
          (target.sheetId && assertSheetUnprotected(target.sheetId),
          operation.kind === "delete")
        )
          throw Error(
            `This name is referenced by ${target.label}. Remove or replace its dependencies before deletion.`,
          );
      }
      const after = replaceTextSpans(before, beforeReferences, (span) => {
        if (operation.kind !== "rename") return span.text;
        const originalDefinitions = sheetIds.map((lookupSheetId) =>
            findScopedDefinedName(span, workbook, lookupSheetId),
          ),
          matchesTarget = originalDefinitions.map(
            (originalDefinition) =>
              !!originalDefinition &&
              definedNameScopeKey(originalDefinition) === existingScopeKey,
          );
        if (matchesTarget.some(Boolean) && !matchesTarget.every(Boolean))
          throw Error(
            `The workbook-scoped definition ${target.label} resolves differently across worksheets. Qualify its references before renaming.`,
          );
        if (matchesTarget.every(Boolean)) {
          if (
            !span.qualifier &&
            span.bound?.has(newDefinition.name.toLowerCase())
          )
            throw Error(
              `Renaming would be shadowed by a local LET variable or LAMBDA parameter in ${target.label}. Choose a distinct name.`,
            );
          const renamedReference = { ...span, name: newDefinition.name };
          if (
            sheetIds.some((checkSheetId) => {
              const scopedDefinedName = findScopedDefinedName(
                renamedReference,
                next,
                checkSheetId,
              );
              return (
                !scopedDefinedName ||
                definedNameScopeKey(scopedDefinedName) !==
                  definedNameScopeKey(newDefinition)
              );
            })
          )
            throw Error(
              `Renaming would be shadowed by another name in ${target.label}. Choose a distinct name.`,
            );
          return span.qualifier
            ? span.text.slice(0, span.text.length - span.name.length) +
                newDefinition.name
            : newDefinition.name;
        }
        if (
          sheetIds.some((checkedSheetId, index) => {
            const scopedDefinedName = findScopedDefinedName(
              span,
              next,
              checkedSheetId,
            );
            return (
              (originalDefinitions[index]
                ? definedNameScopeKey(originalDefinitions[index])
                : undefined) !==
              (scopedDefinedName
                ? definedNameScopeKey(scopedDefinedName)
                : undefined)
            );
          })
        )
          throw Error(
            `Renaming would change an existing reference in ${target.label}. Choose a distinct name.`,
          );
        return span.text;
      });
      if (
        after !== before &&
        (target.sheetId && assertSheetUnprotected(target.sheetId),
        changes.push({ ...target, before: before, after: after }),
        changes.length > pO)
      )
        throw Error(
          "A name operation may rewrite at most 500 formulas or definitions. Split the migration into smaller workbooks.",
        );
      return after;
    };
  for (const sheet of next.sheets) {
    for (const sheetTable of sheet.tables ?? [])
      for (const column of sheetTable.columns)
        column.calculatedFormula &&= rewriteFormula(
          column.calculatedFormula,
          [sheet.id],
          {
            kind: "table-column",
            label: `${sheet.name}: ${sheetTable.name}[${column.name}] calculated rule`,
            sheetId: sheet.id,
          },
        );
    for (const [address, cell] of Object.entries(sheet.cells))
      typeof cell.raw == "string" &&
        cell.raw.startsWith("=") &&
        !cell.literal &&
        (cell.raw = rewriteFormula(cell.raw, [sheet.id], {
          kind: "cell",
          label: `${sheet.name}!${address}`,
          sheetId: sheet.id,
          address: address,
        }));
    for (const rule of [
      ...(sheet.conditionalFormats ?? []),
      ...(sheet.validations ?? []),
    ])
      (rule.kind === "formula" || rule.kind === "custom") &&
        rule.formula &&
        (rule.formula = rewriteFormula(rule.formula, [sheet.id], {
          kind: rule.kind === "custom" ? "validation" : "conditional",
          label: `${sheet.name}: ${rule.kind === "custom" ? "validation" : "conditional"} rule ${rule.id}`,
          sheetId: sheet.id,
        }));
  }
  for (const definedName of next.definedNames) {
    const targetSheetIds = definedName.scopeSheetId
      ? [definedName.scopeSheetId]
      : workbook.sheets.map((sheet) => sheet.id);
    definedName.refersTo = rewriteFormula(
      definedName.refersTo,
      targetSheetIds,
      {
        kind: "name",
        label: `Name ${definedName.name} (${definedName.scopeSheetId ? workbook.sheets.find((sheet) => sheet.id === definedName.scopeSheetId).name : "Workbook"})`,
        ...(definedName.scopeSheetId
          ? { sheetId: definedName.scopeSheetId }
          : {}),
      },
    );
  }
  if (newDefinition)
    for (const targetSheetId of newDefinition.scopeSheetId
      ? [newDefinition.scopeSheetId]
      : workbook.sheets.map((sheet) => sheet.id))
      try {
        resolveDefinedNameValue(
          next.definedNames
            .find(
              (definedName) =>
                definedNameScopeKey(definedName) ===
                definedNameScopeKey(newDefinition),
            )
            .refersTo.slice(1),
          next,
          targetSheetId,
          [definedNameScopeKey(newDefinition)],
        );
      } catch (error) {
        throw Error(
          `The resulting name has a cycle or exceeds the expansion limit: ${error instanceof Error ? error.message : String(error)}`,
        );
      }
  return {
    snapshot: JSON.stringify(workbook),
    request: structuredClone(operation),
    ...(existingName ? { before: structuredClone(existingName) } : {}),
    ...(newDefinition
      ? {
          after: structuredClone(
            next.definedNames.find(
              (definedName) =>
                definedNameScopeKey(definedName) ===
                definedNameScopeKey(newDefinition),
            ),
          ),
        }
      : {}),
    changes: changes,
    dependencies: dependencies,
    next: next,
  };
}
function assertNamePreviewCurrent(workbook, preview, options = {}) {
  if (JSON.stringify(workbook) !== preview.snapshot)
    throw Error(
      "The workbook changed after the name preview. Review the operation again.",
    );
  const recomputedPreview = applyDefinedNameOperation(
    workbook,
    preview.request,
    options,
  );
  if (JSON.stringify(recomputedPreview) !== JSON.stringify(preview))
    throw Error("The name preview changed. Review the operation again.");
  return JSON.stringify(recomputedPreview.next) === JSON.stringify(workbook)
    ? workbook
    : recomputedPreview.next;
}
function matchCriteriaValue(cellValue, criterion) {
  if (typeof criterion != "string") return cellValue === criterion;
  const criterionMatch = /^(<=|>=|<>|=|<|>)?(.*)$/.exec(criterion),
    operator = criterionMatch[1] ?? "=",
    operand = criterionMatch[2];
  if ((operator === "=" || operator === "<>") && /[?*~]/.test(operand)) {
    if (typeof cellValue != "string") return operator === "<>";
    const patternTokens = [];
    for (let patternIndex = 0; patternIndex < operand.length; patternIndex++)
      operand[patternIndex] === "~" && patternIndex + 1 < operand.length
        ? patternTokens.push({
            char: operand[++patternIndex].toLowerCase(),
            wildcard: false,
          })
        : patternTokens.push({
            char: operand[patternIndex].toLowerCase(),
            wildcard:
              operand[patternIndex] === "*" || operand[patternIndex] === "?",
          });
    const cellText = String(cellValue ?? "").toLowerCase();
    if (patternTokens.length * cellText.length > 1e5)
      throw formulaErrorFactory("#LIMIT!");
    let matchRow = Array(cellText.length + 1).fill(false);
    matchRow[0] = true;
    for (const patternToken of patternTokens) {
      const nextMatchRow = Array(cellText.length + 1).fill(false);
      nextMatchRow[0] =
        patternToken.wildcard && patternToken.char === "*" && matchRow[0];
      for (let textIndex = 1; textIndex <= cellText.length; textIndex++)
        nextMatchRow[textIndex] =
          patternToken.wildcard && patternToken.char === "*"
            ? matchRow[textIndex] || nextMatchRow[textIndex - 1]
            : matchRow[textIndex - 1] &&
              ((patternToken.wildcard && patternToken.char === "?") ||
                cellText[textIndex - 1] === patternToken.char);
      matchRow = nextMatchRow;
    }
    return operator === "<>"
      ? !matchRow[cellText.length]
      : matchRow[cellText.length];
  }
  const isNumericOperand =
      operand.trim() !== "" && Number.isFinite(Number(operand)),
    comparableValue = isNumericOperand
      ? cellValue === null || cellValue === ""
        ? 0
        : typeof cellValue == "number"
          ? cellValue
          : NaN
      : String(cellValue ?? "").toLowerCase(),
    comparableOperand = isNumericOperand
      ? Number(operand)
      : operand.toLowerCase();
  return operator === "<>"
    ? comparableValue !== comparableOperand
    : operator === "<"
      ? comparableValue < comparableOperand
      : operator === ">"
        ? comparableValue > comparableOperand
        : operator === "<="
          ? comparableValue <= comparableOperand
          : operator === ">="
            ? comparableValue >= comparableOperand
            : comparableValue === comparableOperand;
}
function createWorkbookFormulaEngine(
  engine,
  referenceFunctionEvaluator,
  issue = createFormulaError,
  clock = resolveLocalDateTime(),
  report,
  criterionFormula,
) {
  const originalApplyInfix = engine.utils._applyInfix.bind(engine.utils);
  ((engine.utils._applyInfix = (leftOperand, operator, rightOperand) => {
    const applyScalarInfix = (val, rightValue) => (
      ["=", "<>", "<", ">", "<=", ">="].includes(operator) &&
        typeof val == "string" &&
        typeof rightValue == "string" &&
        ((val = val.toLowerCase()), (rightValue = rightValue.toLowerCase())),
      originalApplyInfix({ val: val, isArray: false }, operator, {
        val: rightValue,
        isArray: false,
      })
    );
    if (!Array.isArray(leftOperand.val) && !Array.isArray(rightOperand.val))
      return applyScalarInfix(leftOperand.val, rightOperand.val);
    const toMatrix = (operand) =>
        Array.isArray(operand)
          ? operand.every(Array.isArray)
            ? operand
            : [operand]
          : [[operand]],
      leftMatrix = toMatrix(leftOperand.val),
      rightMatrix = toMatrix(rightOperand.val),
      rowCount = Math.max(leftMatrix.length, rightMatrix.length),
      columnCount = Math.max(
        leftMatrix[0]?.length ?? 0,
        rightMatrix[0]?.length ?? 0,
      );
    return !rowCount ||
      !columnCount ||
      (leftMatrix.length !== rowCount && leftMatrix.length !== 1) ||
      (rightMatrix.length !== rowCount && rightMatrix.length !== 1) ||
      (leftMatrix[0].length !== columnCount && leftMatrix[0].length !== 1) ||
      (rightMatrix[0].length !== columnCount && rightMatrix[0].length !== 1)
      ? issue("#VALUE!")
      : rowCount * columnCount > fileImportLimits.maxSpillCells
        ? issue("#LIMIT!")
        : Array.from({ length: rowCount }, (w, rowIndex) =>
            Array.from({ length: columnCount }, (A, columnIndex) =>
              applyScalarInfix(
                leftMatrix[leftMatrix.length === 1 ? 0 : rowIndex][
                  leftMatrix[0].length === 1 ? 0 : columnIndex
                ],
                rightMatrix[rightMatrix.length === 1 ? 0 : rowIndex][
                  rightMatrix[0].length === 1 ? 0 : columnIndex
                ],
              ),
            ),
          );
  }),
    (engine.checkFormulaResult = (reference, allowArrayResult) => {
      const retrievedValue = engine.retrieveRef(reference);
      return isFormulaError(retrievedValue)
        ? retrievedValue
        : Array.isArray(retrievedValue)
          ? allowArrayResult
            ? retrievedValue
            : issue("#VALUE!")
          : retrievedValue === null
            ? 0
            : typeof retrievedValue == "number"
              ? Number.isFinite(retrievedValue)
                ? retrievedValue
                : issue("#NUM!")
              : typeof retrievedValue == "string" ||
                  typeof retrievedValue == "boolean"
                ? retrievedValue
                : issue("#VALUE!");
    }));
  const distribution = engine._callFunction.bind(engine);
  return (
    (engine._callFunction = (functionName, args) => {
      functionName = functionName
        .replace(/^_xlfn\./i, "")
        .replace(/^_xlws\./i, "")
        .toUpperCase();
      const retrieve = (reference) =>
        reference == null ? undefined : engine.retrieveRef(reference);
      try {
        if (!builtinFunctionNameSet.has(functionName)) return issue("#NAME?");
        if (functionName === "RAND")
          return args.length ? issue("#VALUE!") : Math.random();
        if (functionName === "RANDBETWEEN") {
          if (args.length !== 2) return issue("#VALUE!");
          const retrievedArgs = args.map(retrieve),
            formulaErrorArg = findFormulaErrorArg(retrievedArgs);
          if (formulaErrorArg) return formulaErrorArg;
          if (
            retrievedArgs.some(
              (arg) =>
                Array.isArray(arg) ||
                arg === undefined ||
                (typeof arg == "string" && arg.trim() === "") ||
                !Number.isFinite(Number(arg)),
            )
          )
            return issue("#VALUE!");
          const lowerBound = Math.ceil(Number(retrievedArgs[0])),
            upperBound = Math.floor(Number(retrievedArgs[1]));
          return !Number.isSafeInteger(lowerBound) ||
            !Number.isSafeInteger(upperBound) ||
            lowerBound > upperBound ||
            !Number.isSafeInteger(upperBound - lowerBound + 1)
            ? issue("#NUM!")
            : lowerBound +
                Math.floor(Math.random() * (upperBound - lowerBound + 1));
        }
        const dateResult = evaluateDateFunction(functionName, args, {
          retrieve: retrieve,
          issue: issue,
          isIssue: isFormulaError,
          clock: clock,
        });
        if (dateResult !== notDateFunctionSentinel) return dateResult;
        const financialResult = evaluateFinancialFunction(functionName, args, {
          retrieve: retrieve,
          issue: issue,
          isIssue: isFormulaError,
          clock: clock,
          report: report,
        });
        if (financialResult !== notFinancialSentinel) return financialResult;
        const referenceResult = referenceFunctionEvaluator(
          functionName,
          args,
          engine,
        );
        if (referenceResult !== notReferenceFunctionMarker)
          return referenceResult;
        const databaseResult = evaluateDatabaseFunction(functionName, args, {
          retrieve: retrieve,
          issue: issue,
          isIssue: isFormulaError,
          criterionFormula: criterionFormula,
        });
        if (databaseResult !== notDatabaseFunctionMarker) return databaseResult;
        const specialistResult = evaluateSpecialistFunction(
          functionName,
          args,
          {
            retrieve: retrieve,
            issue: issue,
            isIssue: isFormulaError,
          },
        );
        if (specialistResult !== notSpecialistSentinel) return specialistResult;
        const regressionResult = evaluateRegressionFunction(
          functionName,
          args,
          {
            retrieve: retrieve,
            issue: issue,
            isIssue: isFormulaError,
          },
        );
        if (regressionResult !== notRegressionSentinel) return regressionResult;
        const significanceResult = evaluateSignificanceTestFunction(
          functionName,
          args,
          {
            retrieve: retrieve,
            issue: issue,
            isIssue: isFormulaError,
            distribution: distribution,
          },
        );
        if (significanceResult !== notInferenceSentinel)
          return significanceResult;
        const arithmeticResult = evaluateArithmeticFunction(
          functionName,
          args,
          {
            retrieve: retrieve,
            issue: issue,
            isIssue: isFormulaError,
          },
        );
        if (arithmeticResult !== notArithmeticSentinel) return arithmeticResult;
        const statisticalResult = evaluateStatisticalFunction(
          functionName,
          args,
          {
            retrieve: retrieve,
            issue: issue,
            isIssue: isFormulaError,
          },
        );
        if (statisticalResult !== notStatisticalSentinel)
          return statisticalResult;
        const logicalResult = evaluateLogicalFunction(functionName, args, {
          retrieve: retrieve,
          issue: issue,
          isIssue: isFormulaError,
        });
        if (logicalResult !== notOwnedLogicalSentinel) return logicalResult;
        const lookupResult = evaluateLookupFunction(functionName, args, {
          retrieve: retrieve,
          issue: issue,
          isIssue: isFormulaError,
          matches: matchCriteriaValue,
        });
        if (lookupResult !== notOwnedLookupSentinel) return lookupResult;
        if (functionName === "TEXTBEFORE" || functionName === "TEXTAFTER")
          return broadcastOverArrays(
            args.map(retrieve),
            (argValues) => {
              try {
                return evaluateTextFunction(functionName, argValues, {
                  retrieve: (value) => value,
                  issue: issue,
                  isIssue: isFormulaError,
                  format: () => issue("#VALUE!"),
                });
              } catch (error) {
                return isFormulaError(error) ? error : issue("#VALUE!");
              }
            },
            issue,
          );
        const textResult = evaluateTextFunction(functionName, args, {
          retrieve: retrieve,
          issue: issue,
          isIssue: isFormulaError,
          temporal: (dateText) => parseDateText(dateText, clock),
          format: () => {
            const serialValue = Number(retrieve(args[0]) ?? 0),
              pattern = retrieve(args[1]);
            if (typeof serialValue == "number" && typeof pattern == "string") {
              try {
                if (validateNumberFormatCode(pattern).kind === "temporal")
                  return formatCellValueWithNumberFormat(
                    serialValue,
                    { category: "custom", locale: "en-US", pattern: pattern },
                    clock.dateSystem,
                  ).text;
              } catch {}
              const unquotedPattern = pattern.replace(/"[^"]*"|\\./g, "");
              if (
                clock.dateSystem === "1904" &&
                /[ydm]/i.test(unquotedPattern) &&
                !/\[[hms]+\]/i.test(unquotedPattern)
              )
                return distribution(functionName, [
                  serialValue + dateSystemSerialOffset(clock.dateSystem),
                  pattern,
                ]);
            }
            return distribution(functionName, args);
          },
        });
        if (textResult !== notTextFunctionSentinel) return textResult;
        const definedNameResult = evaluateDefinedNameFunction(
          functionName,
          args,
          {
            retrieve: retrieve,
            issue: issue,
            isIssue: isFormulaError,
            matches: matchCriteriaValue,
          },
        );
        if (definedNameResult !== notModernFunctionSentinel)
          return definedNameResult;
        if (functionName === "INDEX") {
          if (args.length < 2 || args.length > 4) return issue("#VALUE!");
          if (isFormulaError(args[0])) return args[0];
          const indexArgs = [...args];
          for (let argIndex = 1; argIndex < args.length; argIndex++) {
            const argValue =
              args[argIndex] == null
                ? argIndex === 3
                  ? 1
                  : 0
                : retrieve(args[argIndex]);
            if (isFormulaError(argValue)) return argValue;
            if (
              Array.isArray(argValue) ||
              (typeof argValue == "string" && argValue.trim() === "") ||
              !Number.isFinite(Number(argValue)) ||
              Number(argValue) < 0
            )
              return issue("#VALUE!");
            indexArgs[argIndex] = Number(argValue);
          }
          if (indexArgs.length > 3 && Number(indexArgs[3]) < 1)
            return issue("#VALUE!");
          const indexResult = distribution(functionName, indexArgs);
          if (
            indexResult &&
            typeof indexResult == "object" &&
            "ref" in indexResult
          ) {
            const sourceArg = args[0],
              resultToUpdate = indexResult;
            sourceArg?.ref?.sheet &&
              !resultToUpdate.ref.sheet &&
              (resultToUpdate.ref.sheet = sourceArg.ref.sheet);
          }
          return indexResult;
        }
        if (functionName === "IF") {
          if (args.length < 2 || args.length > 3) return issue("#VALUE!");
          const condition = retrieve(args[0]);
          if (isFormulaError(condition)) return condition;
          if (Array.isArray(condition))
            return broadcastOverArrays(
              [
                condition,
                retrieve(args[1]),
                args.length === 3 ? retrieve(args[2]) : false,
              ],
              ([conditionValue, trueValue, falseValue]) =>
                isFormulaError(conditionValue)
                  ? conditionValue
                  : typeof conditionValue == "string" &&
                      !/^(TRUE|FALSE)$/i.test(conditionValue)
                    ? issue("#VALUE!")
                    : (((
                        typeof conditionValue == "string"
                          ? conditionValue.toUpperCase() === "TRUE"
                          : !!conditionValue
                      )
                        ? trueValue
                        : falseValue) ?? 0),
              issue,
            );
          if (
            typeof condition == "string" &&
            !/^(TRUE|FALSE)$/i.test(condition)
          )
            return issue("#VALUE!");
          const isConditionTrue =
            typeof condition == "string"
              ? condition.toUpperCase() === "TRUE"
              : !!condition;
          return (
            retrieve(
              isConditionTrue ? args[1] : args.length === 3 ? args[2] : false,
            ) ?? 0
          );
        }
        if (functionName === "IFERROR" || functionName === "IFNA") {
          if (args.length !== 2) return issue("#VALUE!");
          const primaryValue = retrieve(args[0]);
          return Array.isArray(primaryValue)
            ? broadcastOverArrays(
                [primaryValue, retrieve(args[1])],
                ([mainValue, fallbackValue]) =>
                  isFormulaError(mainValue) &&
                  (functionName === "IFERROR" || mainValue.error === "#N/A")
                    ? (fallbackValue ?? 0)
                    : (mainValue ?? 0),
                issue,
              )
            : isFormulaError(primaryValue) &&
                (functionName === "IFERROR" || primaryValue.error === "#N/A")
              ? (retrieve(args[1]) ?? 0)
              : (primaryValue ?? 0);
        }
        const resolvedArgs = args.map(retrieve),
          numericArgs = resolvedArgs
            .flatMap(flattenNestedValues)
            .filter((entry) => typeof entry == "number");
        if (
          resolvedArgs.flatMap(flattenNestedValues).length > 1e5 ||
          ([
            "REPT",
            "BASE",
            "DEC2BIN",
            "DEC2HEX",
            "DEC2OCT",
            "BIN2HEX",
            "BIN2OCT",
            "HEX2BIN",
            "HEX2OCT",
            "OCT2BIN",
            "OCT2HEX",
            "FIXED",
            "DOLLAR",
          ].includes(functionName) &&
            ((Math.abs(
              Number(resolvedArgs[functionName === "BASE" ? 2 : 1] ?? 0),
            ) > 32767 &&
              !["FIXED", "DOLLAR"].includes(functionName)) ||
              (["FIXED", "DOLLAR"].includes(functionName) &&
                Math.abs(Number(resolvedArgs[1] ?? 2)) > 100) ||
              (functionName === "REPT" &&
                String(resolvedArgs[0] ?? "").length * Number(resolvedArgs[1]) >
                  32767))) ||
          (["MUNIT", "MDETERM", "MMULT", "GROWTH"].includes(functionName) &&
            (functionName === "MUNIT"
              ? Number(resolvedArgs[0]) > 100
              : resolvedArgs.flatMap(flattenNestedValues).length > 1e3)) ||
          ([
            "BINOM.DIST.RANGE",
            "BINOM.INV",
            "HYPGEOM.DIST",
            "POISSON.DIST",
            "FACT",
            "FACTDOUBLE",
            "COMBIN",
            "COMBINA",
            "MULTINOMIAL",
          ].includes(functionName) &&
            numericArgs.some((numericValue) => Math.abs(numericValue) > 1e4)) ||
          (["BESSELI", "BESSELJ", "BESSELK", "BESSELY"].includes(
            functionName,
          ) &&
            Math.abs(Number(resolvedArgs[1])) > 1e3)
        )
          return issue("#LIMIT!");
        if (
          functionName === "ROMAN" &&
          Math.abs(Number(resolvedArgs[0])) > 3999
        )
          return issue("#VALUE!");
        if (
          [
            "NETWORKDAYS",
            "NETWORKDAYS.INTL",
            "WORKDAY",
            "WORKDAY.INTL",
          ].includes(functionName) &&
          (Math.abs(
            Number(resolvedArgs[1]) -
              (functionName.startsWith("NETWORK")
                ? Number(resolvedArgs[0])
                : 0),
          ) > 366e3 ||
            resolvedArgs.flatMap(flattenNestedValues).length > 1e3)
        )
          return issue("#LIMIT!");
        if (functionName === "ISERR")
          return (
            isFormulaError(resolvedArgs[0]) && resolvedArgs[0].error !== "#N/A"
          );
        if (functionName === "ISLOGICAL")
          return typeof resolvedArgs[0] == "boolean";
        if (functionName === "ISNONTEXT")
          return typeof resolvedArgs[0] != "string";
        if (functionName === "N")
          return isFormulaError(resolvedArgs[0]) ||
            typeof resolvedArgs[0] == "number"
            ? resolvedArgs[0]
            : resolvedArgs[0] === true
              ? 1
              : 0;
        if (functionName === "T")
          return isFormulaError(resolvedArgs[0]) ||
            typeof resolvedArgs[0] == "string"
            ? resolvedArgs[0]
            : "";
        if (functionName === "TYPE")
          return isFormulaError(resolvedArgs[0])
            ? 16
            : Array.isArray(resolvedArgs[0])
              ? 64
              : typeof resolvedArgs[0] == "string"
                ? 2
                : typeof resolvedArgs[0] == "boolean"
                  ? 4
                  : 1;
        if (functionName === "ERROR.TYPE") {
          const errorIndex = [
            "#NULL!",
            "#DIV/0!",
            "#VALUE!",
            "#REF!",
            "#NAME?",
            "#NUM!",
            "#N/A",
            "#GETTING_DATA",
          ].indexOf(
            isFormulaError(resolvedArgs[0]) ? resolvedArgs[0].error : "",
          );
          return errorIndex < 0 ? issue("#N/A") : errorIndex + 1;
        }
        if (functionName === "ISERROR") return isFormulaError(resolvedArgs[0]);
        if (functionName === "ISNA")
          return (
            isFormulaError(resolvedArgs[0]) && resolvedArgs[0].error === "#N/A"
          );
        if (functionName === "ISNUMBER")
          return typeof resolvedArgs[0] == "number";
        if (functionName === "ISTEXT")
          return typeof resolvedArgs[0] == "string";
        if (functionName === "ISBLANK") return resolvedArgs[0] === null;
        if (functionName === "COUNTA")
          return resolvedArgs
            .flatMap(flattenNestedValues)
            .filter((entry) => entry !== null).length;
        if (functionName === "COUNTBLANK")
          return resolvedArgs
            .flatMap(flattenNestedValues)
            .filter((entry) => entry === null || entry === "").length;
        if (
          [
            "SUMIF",
            "AVERAGEIF",
            "COUNTIF",
            "SUMIFS",
            "AVERAGEIFS",
            "COUNTIFS",
          ].includes(functionName)
        ) {
          const isMultiCriteria = functionName.endsWith("IFS"),
            isCountFunction = functionName.startsWith("COUNT"),
            isAverageFunction = functionName.startsWith("AVERAGE");
          let valueCells, criteriaArgs;
          if (isMultiCriteria) {
            if (
              ((criteriaArgs = isCountFunction
                ? resolvedArgs
                : resolvedArgs.slice(1)),
              criteriaArgs.length < 2 || criteriaArgs.length % 2 !== 0)
            )
              return issue("#VALUE!");
            valueCells = isCountFunction
              ? flattenNestedValues(criteriaArgs[0]).map(() => 1)
              : flattenNestedValues(resolvedArgs[0]);
          } else {
            if (
              resolvedArgs.length < 2 ||
              resolvedArgs.length > (isCountFunction ? 2 : 3)
            )
              return issue("#VALUE!");
            ((criteriaArgs = resolvedArgs.slice(0, 2)),
              (valueCells = isCountFunction
                ? flattenNestedValues(resolvedArgs[0]).map(() => 1)
                : flattenNestedValues(
                    args.length < 3 || args[2] == null
                      ? resolvedArgs[0]
                      : resolvedArgs[2],
                  )));
          }
          const criteriaRanges = [],
            criteriaValues = [],
            rangeShape = (range) => {
              if (!Array.isArray(range)) return "1:1";
              const rows = range.every(Array.isArray) ? range : [range],
                columnCount = rows[0]?.length ?? 0;
              return rows.length &&
                columnCount &&
                rows.every((row) => row.length === columnCount)
                ? `${rows.length}:${columnCount}`
                : "invalid";
            },
            criteriaShape = rangeShape(
              isCountFunction
                ? criteriaArgs[0]
                : isMultiCriteria || args.length < 3 || args[2] == null
                  ? resolvedArgs[0]
                  : resolvedArgs[2],
            );
          for (
            let pairIndex = 0;
            pairIndex < criteriaArgs.length;
            pairIndex += 2
          ) {
            const criteriaCells = flattenNestedValues(criteriaArgs[pairIndex]);
            if (
              criteriaCells.length !== valueCells.length ||
              criteriaShape === "invalid" ||
              rangeShape(criteriaArgs[pairIndex]) !== criteriaShape ||
              Array.isArray(criteriaArgs[pairIndex + 1])
            )
              return issue("#VALUE!");
            if (isFormulaError(criteriaArgs[pairIndex + 1]))
              return criteriaArgs[pairIndex + 1];
            criteriaRanges.push(criteriaCells);
            const criterion = criteriaArgs[pairIndex + 1];
            criteriaValues.push(
              criterion == null
                ? "=0"
                : typeof criterion == "number"
                  ? `=${criterion}`
                  : criterion,
            );
          }
          const matchedValues = [];
          let workUnits = 0;
          for (let cellIndex = 0; cellIndex < valueCells.length; cellIndex++) {
            if (
              ((workUnits += criteriaRanges.reduce(
                (total, rangeCells, rangeIndex) =>
                  total +
                  Math.max(1, String(rangeCells[cellIndex] ?? "").length) *
                    Math.max(
                      1,
                      String(criteriaValues[rangeIndex] ?? "").length,
                    ),
                0,
              )),
              workUnits > 2e6)
            )
              return issue("#LIMIT!");
            if (
              criteriaRanges.every(
                (rangeCells, rangeIndex) =>
                  !isFormulaError(rangeCells[cellIndex]) &&
                  matchCriteriaValue(
                    rangeCells[cellIndex],
                    criteriaValues[rangeIndex],
                  ),
              )
            ) {
              if (isFormulaError(valueCells[cellIndex]))
                return valueCells[cellIndex];
              typeof valueCells[cellIndex] == "number" &&
                matchedValues.push(valueCells[cellIndex]);
            }
          }
          return isAverageFunction
            ? matchedValues.length
              ? exactAverageOfNumbers(matchedValues)
              : issue("#DIV/0!")
            : sumDecimalPartsToNumber(matchedValues);
        }
        const formulaErrorArg = findFormulaErrorArg(resolvedArgs);
        if (formulaErrorArg && functionName !== "COUNT") return formulaErrorArg;
        if (functionName === "TRIM")
          return String(resolvedArgs[0] ?? "")
            .replace(/^ +| +$/g, "")
            .replace(/ +/g, " ");
        if (functionName === "UPPER") {
          const upperText = String(resolvedArgs[0] ?? "").toUpperCase();
          return upperText.length > 32767 ? issue("#LIMIT!") : upperText;
        }
        if (functionName === "VALUE") {
          if (resolvedArgs.length !== 1 || Array.isArray(resolvedArgs[0]))
            return issue("#VALUE!");
          let numberText = String(resolvedArgs[0] ?? "").trim(),
            isNegative = false;
          numberText.startsWith("(") &&
            numberText.endsWith(")") &&
            ((isNegative = true),
            (numberText = numberText.slice(1, -1).trim()));
          const isPercent = numberText.endsWith("%");
          if (
            (isPercent && (numberText = numberText.slice(0, -1).trim()),
            (numberText = numberText.replace(/^\$/, "")),
            !/^[+-]?(?:\d{1,3}(?:,\d{3})+|\d+|\d*\.\d+)(?:\.\d*)?(?:[eE][+-]?\d+)?$/.test(
              numberText,
            ))
          )
            return issue("#VALUE!");
          const parsedNumber =
            (Number(numberText.replace(/,/g, "")) * (isNegative ? -1 : 1)) /
            (isPercent ? 100 : 1);
          return Number.isFinite(parsedNumber) ? parsedNumber : issue("#NUM!");
        }
        if (functionName === "MATCH") {
          if (resolvedArgs.length < 2 || resolvedArgs.length > 3)
            return issue("#VALUE!");
          const lookupArray = resolvedArgs[1];
          if (
            Array.isArray(lookupArray) &&
            lookupArray.length > 1 &&
            Array.isArray(lookupArray[0]) &&
            lookupArray[0].length > 1
          )
            return issue("#N/A");
          const lookupValues = flattenNestedValues(lookupArray),
            lookupValue = resolvedArgs[0],
            matchType =
              resolvedArgs[2] === undefined ? 1 : Number(resolvedArgs[2]);
          if (![0, -1, 1].includes(matchType)) return issue("#N/A");
          if (matchType === 0) {
            const matchIndex = lookupValues.findIndex((candidate) =>
              typeof lookupValue == "string"
                ? typeof candidate == "string" &&
                  (/[?*~]/.test(lookupValue)
                    ? matchCriteriaValue(candidate, `=${lookupValue}`)
                    : candidate.toLowerCase() === lookupValue.toLowerCase())
                : candidate === lookupValue,
            );
            return matchIndex < 0 ? issue("#N/A") : matchIndex + 1;
          }
          const normalizedValues = lookupValues.map((lookupItem) =>
              typeof lookupItem == "string"
                ? lookupItem.toLowerCase()
                : lookupItem,
            ),
            normalizedLookup =
              typeof lookupValue == "string"
                ? lookupValue.toLowerCase()
                : lookupValue;
          if (
            (typeof normalizedLookup != "number" &&
              typeof normalizedLookup != "string") ||
            normalizedValues.some(
              (normalizedValue) =>
                typeof normalizedValue != typeof normalizedLookup,
            )
          )
            return issue("#VALUE!");
          const sortedValues = normalizedValues;
          if (
            sortedValues.some(
              (sortedValue, valueIndex) =>
                valueIndex > 0 &&
                (matchType === 1
                  ? sortedValue < sortedValues[valueIndex - 1]
                  : sortedValue > sortedValues[valueIndex - 1]),
            )
          )
            return issue("#VALUE!");
          let lastMatchIndex = -1;
          return (
            sortedValues.forEach((sortedValue, valueIndex) => {
              (matchType === 1
                ? sortedValue <= normalizedLookup
                : sortedValue >= normalizedLookup) &&
                (lastMatchIndex = valueIndex);
            }),
            lastMatchIndex < 0 ? issue("#N/A") : lastMatchIndex + 1
          );
        }
        if (
          [
            "SUM",
            "AVERAGE",
            "MIN",
            "MAX",
            "MEDIAN",
            "LARGE",
            "SMALL",
            "STDEV.S",
            "VAR.S",
          ].includes(functionName)
        ) {
          const valueArgs =
              functionName === "LARGE" || functionName === "SMALL"
                ? resolvedArgs.slice(0, 1)
                : resolvedArgs,
            numericValues = [];
          valueArgs.forEach((valueArg, argIndex) => {
            const isReferenceArg =
              !!args[argIndex] &&
              typeof args[argIndex] == "object" &&
              "ref" in args[argIndex];
            for (const flatValue of flattenNestedValues(valueArg))
              typeof flatValue == "number"
                ? numericValues.push(flatValue)
                : !isReferenceArg &&
                  !Array.isArray(valueArg) &&
                  (typeof flatValue == "boolean" ||
                    (typeof flatValue == "string" &&
                      flatValue.trim() !== "" &&
                      Number.isFinite(Number(flatValue)))) &&
                  numericValues.push(Number(flatValue));
          });
          const sum = sumDecimalPartsToNumber(numericValues);
          if (functionName === "SUM") return sum;
          if (functionName === "AVERAGE")
            return numericValues.length
              ? sum / numericValues.length
              : issue("#DIV/0!");
          if (functionName === "MIN")
            return numericValues.length
              ? numericValues.reduce((minimum, candidate) =>
                  Math.min(minimum, candidate),
                )
              : 0;
          if (functionName === "MAX")
            return numericValues.length
              ? numericValues.reduce((maximum, candidate) =>
                  Math.max(maximum, candidate),
                )
              : 0;
          if (functionName === "STDEV.S" || functionName === "VAR.S") {
            if (numericValues.length < 2) return issue("#DIV/0!");
            const variance =
              numericValues.reduce(
                (total, numericValue) =>
                  total + (numericValue - sum / numericValues.length) ** 2,
                0,
              ) /
              (numericValues.length - 1);
            return functionName === "VAR.S" ? variance : Math.sqrt(variance);
          }
          if (
            (numericValues.sort(
              (leftNumber, rightNumber) => leftNumber - rightNumber,
            ),
            functionName === "MEDIAN")
          )
            return numericValues.length
              ? (numericValues[Math.floor((numericValues.length - 1) / 2)] +
                  numericValues[Math.floor(numericValues.length / 2)]) /
                  2
              : issue("#NUM!");
          const position = Number(resolvedArgs[1]);
          return !Number.isInteger(position) ||
            position < 1 ||
            position > numericValues.length
            ? issue("#NUM!")
            : functionName === "LARGE"
              ? numericValues[numericValues.length - position]
              : numericValues[position - 1];
        }
        if (functionName === "SUBSTITUTE") {
          if (
            resolvedArgs.length < 3 ||
            resolvedArgs.length > 4 ||
            resolvedArgs.some(Array.isArray)
          )
            return issue("#VALUE!");
          const sourceText = String(resolvedArgs[0] ?? ""),
            oldText = String(resolvedArgs[1] ?? ""),
            newText = String(resolvedArgs[2] ?? ""),
            occurrenceCount = oldText.length
              ? sourceText.split(oldText).length - 1
              : 0;
          if (
            resolvedArgs[3] !== undefined &&
            (!Number.isInteger(Number(resolvedArgs[3])) ||
              Number(resolvedArgs[3]) < 1)
          )
            return issue("#VALUE!");
          if (
            sourceText.length +
              (resolvedArgs[3] === undefined
                ? occurrenceCount
                : Math.min(occurrenceCount, 1)) *
                (newText.length - oldText.length) >
            32767
          )
            return issue("#LIMIT!");
          if (!oldText.length) return sourceText;
          if (resolvedArgs[3] === undefined)
            return sourceText.split(oldText).join(newText);
          let replacedCount = 0;
          return sourceText
            .split(oldText)
            .reduce(
              (resultText, segment, segmentIndex) =>
                segmentIndex === 0
                  ? segment
                  : resultText +
                    (++replacedCount === Number(resolvedArgs[3])
                      ? newText
                      : oldText) +
                    segment,
              "",
            );
        }
        if (functionName === "SEARCH" || functionName === "SEARCHB") {
          if (
            resolvedArgs.length < 2 ||
            resolvedArgs.length > 3 ||
            resolvedArgs.some(Array.isArray)
          )
            return issue("#VALUE!");
          const searchText = String(resolvedArgs[0] ?? "").toLowerCase(),
            withinText = String(resolvedArgs[1] ?? "").toLowerCase(),
            startPosition =
              resolvedArgs[2] === undefined ? 1 : Number(resolvedArgs[2]);
          if (
            !Number.isInteger(startPosition) ||
            startPosition < 1 ||
            startPosition > withinText.length
          )
            return issue("#VALUE!");
          if (!/[?*~]/.test(searchText)) {
            const foundIndex = withinText.indexOf(
              searchText,
              startPosition - 1,
            );
            return foundIndex < 0 ? issue("#VALUE!") : foundIndex + 1;
          }
          if (withinText.length * withinText.length * searchText.length > 2e6)
            return issue("#LIMIT!");
          for (
            let offset = startPosition - 1;
            offset < withinText.length;
            offset++
          )
            if (matchCriteriaValue(withinText.slice(offset), `${searchText}*`))
              return offset + 1;
          return issue("#VALUE!");
        }
        if (
          functionName === "CONCATENATE" &&
          resolvedArgs
            .flatMap(flattenNestedValues)
            .reduce((total, entry) => total + String(entry ?? "").length, 0) >
            32767
        )
          return issue("#LIMIT!");
        const distributedResult = distribution(functionName, args);
        return /^IM/.test(functionName) &&
          !["IMABS", "IMREAL", "IMAGINARY", "IMARGUMENT"].includes(
            functionName,
          ) &&
          typeof distributedResult == "number"
          ? String(distributedResult)
          : typeof distributedResult == "string" &&
              distributedResult.length > 32767
            ? issue("#LIMIT!")
            : distributedResult;
      } catch (error) {
        return isFormulaError(error) ? issue(error.error) : issue("#ERROR!");
      }
    }),
    engine
  );
}
function formatNumberWithPattern(value, cell, dateSystem = "1900") {
  if (cell?.numberFormat)
    try {
      return formatCellValueWithNumberFormat(
        value,
        cell.numberFormat,
        dateSystem,
      ).text;
    } catch {
      return "#FORMAT!";
    }
  if (value === null) return "";
  if (typeof value == "boolean") return value ? "TRUE" : "FALSE";
  if (typeof value != "number") return value;
  if (!Number.isFinite(value)) return "#NUM!";
  if (cell?.format === "date") {
    if (cell.numFmt)
      try {
        if (validateNumberFormatCode(cell.numFmt).kind === "temporal")
          return formatCellValueWithNumberFormat(
            value,
            { category: "custom", locale: "en-US", pattern: cell.numFmt },
            dateSystem,
          ).text;
      } catch {}
    return formatCellValueWithNumberFormat(
      value,
      { category: "date", locale: "en-US", pattern: "yyyy-mm-dd" },
      dateSystem,
    ).text;
  }
  if (cell?.numFmt) {
    if (cell.numFmt.toLowerCase() === "general")
      return String(Number(value.toPrecision(15)));
    const intlOptions = numberFormatToIntlOptions(value, cell.numFmt);
    if (intlOptions !== undefined) return intlOptions;
  }
  return cell?.format === "percent"
    ? new Intl.NumberFormat("en-US", {
        style: "percent",
        minimumFractionDigits: 2,
        maximumFractionDigits: 2,
      }).format(value)
    : cell?.format === "currency"
      ? new Intl.NumberFormat("en-US", {
          style: "currency",
          currency: "USD",
          minimumFractionDigits: 2,
          maximumFractionDigits: 2,
        }).format(value)
      : cell?.format === "number"
        ? new Intl.NumberFormat("en-US", {
            minimumFractionDigits: 2,
            maximumFractionDigits: 2,
          }).format(value)
        : String(Number(value.toPrecision(15)));
}
const defaultCalculationSettings = Object.freeze({
    mode: "automatic",
    iterative: false,
    maxIterations: 100,
    maxChange: 0.001,
    timeBudgetMs: 2500,
  }),
  throwInvalidCalculationSettings = () => {
    throw Error(
      "Invalid calculation settings. Choose Automatic or Manual, explicit iteration, 1–1000 iterations, convergence from 0 to 1, a 1–10000 ms time budget, 1–128 MiB working data and 1000–10000000 work units.",
    );
  };
function normalizeCalculationSettings(settings) {
  return !isPlainObjectValue(settings) ||
    Object.keys(settings).some(
      (key) =>
        ![
          ...Object.keys(defaultCalculationSettings),
          "memoryBudgetMiB",
          "workBudget",
        ].includes(key),
    ) ||
    typeof settings.mode != "string" ||
    !["automatic", "manual"].includes(settings.mode) ||
    typeof settings.iterative != "boolean" ||
    !Number.isInteger(settings.maxIterations) ||
    Number(settings.maxIterations) < 1 ||
    Number(settings.maxIterations) > 1e3 ||
    typeof settings.maxChange != "number" ||
    !Number.isFinite(settings.maxChange) ||
    settings.maxChange < 0 ||
    settings.maxChange > 1 ||
    !Number.isInteger(settings.timeBudgetMs) ||
    Number(settings.timeBudgetMs) < 1 ||
    Number(settings.timeBudgetMs) > 1e4 ||
    (settings.memoryBudgetMiB !== undefined &&
      (!Number.isInteger(settings.memoryBudgetMiB) ||
        Number(settings.memoryBudgetMiB) < 1 ||
        Number(settings.memoryBudgetMiB) > 128)) ||
    (settings.workBudget !== undefined &&
      (!Number.isInteger(settings.workBudget) ||
        Number(settings.workBudget) < 1e3 ||
        Number(settings.workBudget) > 1e7))
    ? throwInvalidCalculationSettings()
    : {
        mode: settings.mode,
        iterative: settings.iterative,
        maxIterations: Number(settings.maxIterations),
        maxChange: settings.maxChange,
        timeBudgetMs: Number(settings.timeBudgetMs),
        ...(settings.memoryBudgetMiB !== undefined
          ? { memoryBudgetMiB: Number(settings.memoryBudgetMiB) }
          : {}),
        ...(settings.workBudget !== undefined
          ? { workBudget: Number(settings.workBudget) }
          : {}),
      };
}
function importFormulaCacheEvidence(workbook, target) {
  workbook.calculation !== undefined &&
    (target.calculation = normalizeCalculationSettings(workbook.calculation));
  const importedCalculationCaches = workbook.importedCalculationCaches;
  if (importedCalculationCaches === undefined) return;
  const throwInvalidEvidence = () => {
    throw Error(
      "Invalid imported formula-cache evidence. Caches must remain separately labeled source evidence.",
    );
  };
  if (
    !isPlainObjectValue(importedCalculationCaches) ||
    importedCalculationCaches.schemaVersion !== 1 ||
    typeof importedCalculationCaches.fileName != "string" ||
    !importedCalculationCaches.fileName ||
    importedCalculationCaches.fileName.length > 255 ||
    typeof importedCalculationCaches.sourceSha256 != "string" ||
    !/^[a-f0-9]{64}$/.test(importedCalculationCaches.sourceSha256) ||
    typeof importedCalculationCaches.importedAt != "string" ||
    !Number.isFinite(Date.parse(importedCalculationCaches.importedAt)) ||
    !isPlainObjectValue(importedCalculationCaches.sourceSettings) ||
    !Array.isArray(importedCalculationCaches.entries) ||
    importedCalculationCaches.entries.length > 1e5
  )
    return throwInvalidEvidence();
  const sourceSettings = importedCalculationCaches.sourceSettings;
  if (
    typeof sourceSettings.mode != "string" ||
    sourceSettings.mode.length > 40 ||
    typeof sourceSettings.iterative != "boolean" ||
    (sourceSettings.maxIterations !== null &&
      (!Number.isSafeInteger(sourceSettings.maxIterations) ||
        Number(sourceSettings.maxIterations) < 1)) ||
    (sourceSettings.fullPrecision !== undefined &&
      typeof sourceSettings.fullPrecision != "boolean") ||
    (sourceSettings.maxChange !== null &&
      (typeof sourceSettings.maxChange != "number" ||
        !Number.isFinite(sourceSettings.maxChange) ||
        sourceSettings.maxChange < 0))
  )
    return throwInvalidEvidence();
  const isAddressInBounds = (address) => {
      const addressMatch = /^([A-Z]+)([1-9]\d*)$/.exec(address);
      return (
        !!addressMatch &&
        Number(addressMatch[2]) <= 1048576 &&
        [...addressMatch[1]].reduce(
          (total, letter) => total * 26 + letter.charCodeAt(0) - 64,
          0,
        ) <= 16384
      );
    },
    seenEntryKeys = new Set(),
    entries = importedCalculationCaches.entries.map((entry) => {
      if (
        !isPlainObjectValue(entry) ||
        typeof entry.sheetId != "string" ||
        !entry.sheetId ||
        entry.sheetId.length > 100 ||
        typeof entry.address != "string" ||
        !/^[A-Z]{1,3}[1-9]\d{0,6}$/.test(entry.address) ||
        !isAddressInBounds(entry.address) ||
        typeof entry.formula != "string" ||
        !entry.formula.startsWith("=") ||
        entry.formula.length > 8192
      )
        return throwInvalidEvidence();
      const entryKey = JSON.stringify([entry.sheetId, entry.address]);
      if (seenEntryKeys.has(entryKey)) return throwInvalidEvidence();
      seenEntryKeys.add(entryKey);
      let normalizedCache;
      if (entry.cache !== undefined) {
        const cache = entry.cache;
        if (
          !isPlainObjectValue(cache) ||
          !["error", "blank", "number", "text", "boolean"].includes(
            String(cache.type),
          ) ||
          (cache.type === "blank" && cache.value !== null) ||
          (cache.type === "number" &&
            (typeof cache.value != "number" ||
              !Number.isFinite(cache.value))) ||
          (cache.type === "boolean" && typeof cache.value != "boolean") ||
          (["text", "error"].includes(String(cache.type)) &&
            (typeof cache.value != "string" || cache.value.length > 32767)) ||
          (cache.type === "error" &&
            !/^#[A-Z0-9_/?!]+$/.test(String(cache.value)))
        )
          return throwInvalidEvidence();
        normalizedCache = { type: cache.type, value: cache.value };
      }
      return {
        sheetId: entry.sheetId,
        address: entry.address,
        formula: entry.formula,
        ...(normalizedCache ? { cache: normalizedCache } : {}),
      };
    });
  target.importedCalculationCaches = {
    schemaVersion: 1,
    fileName: importedCalculationCaches.fileName,
    sourceSha256: importedCalculationCaches.sourceSha256,
    importedAt: importedCalculationCaches.importedAt,
    sourceSettings: {
      mode: sourceSettings.mode,
      iterative: sourceSettings.iterative,
      maxIterations: sourceSettings.maxIterations,
      maxChange: sourceSettings.maxChange,
      ...(sourceSettings.fullPrecision !== undefined
        ? { fullPrecision: sourceSettings.fullPrecision }
        : {}),
    },
    entries: entries,
  };
}
const calculatorRecords = new WeakMap(),
  carriedCalculatorRecords = new WeakMap(),
  j5 = 2e5,
  pendingFormulaValue = { type: "error", value: "#PENDING!" },
  makeFormulaError = (errorValue) => ({ type: "error", value: errorValue }),
  dynamicFunctionPattern = new RegExp(
    `\\b(?:${[...matrixResultFunctionNames, "LET", "LAMBDA", "INDIRECT", "OFFSET", "ISFORMULA", "TEXTBEFORE", "TEXTAFTER"].join("|")})\\s*\\(`,
    "i",
  ),
  _O = {
    test: (formulaText) => findVolatileFunctionNames(formulaText).length > 0,
  },
  ik = (workbook, formulas) =>
    !!workbook.definedNames?.length ||
    [...formulas.values()].some((formulaEntry) => {
      const formulaWithoutStrings = formulaEntry.formula.replace(
        /"(?:[^"]|"")*"/g,
        '""',
      );
      return (
        dynamicFunctionPattern.test(formulaWithoutStrings) ||
        /[{}#]/.test(formulaWithoutStrings) ||
        /^=\s*[^()+*/]*:/.test(formulaWithoutStrings) ||
        (/\b(?:IF|IFERROR|IFNA)\s*\(/i.test(formulaWithoutStrings) &&
          formulaWithoutStrings.includes(":"))
      );
    }),
  calculationSnapshotCache = new WeakMap();
const sheetInputCache = new WeakMap();
function buildSheetCalculationSnapshot(sheet) {
  const cachedSnapshot = isDeepFrozenValue(sheet.cells)
    ? sheetInputCache.get(sheet.cells)?.get(sheet.id)
    : undefined;
  if (cachedSnapshot) return cachedSnapshot;
  const formulas = [];
  let cachedInputs, cachedSignature;
  const inputs = () =>
      (cachedInputs ??= Object.entries(sheet.cells).map(
        ([cellAddress, cell]) => [
          `${sheet.id}:${cellAddress}`,
          JSON.stringify([cell.raw, !!cell.literal]),
        ],
      )),
    signature = () => (cachedSignature ??= JSON.stringify(inputs()));
  isDeepFrozenValue(sheet.cells) || signature();
  for (const address in sheet.cells) {
    const cell = sheet.cells[address],
      cellKey = `${sheet.id}:${address}`;
    !cell.literal &&
      typeof cell.raw == "string" &&
      cell.raw.startsWith("=") &&
      formulas.push([
        cellKey,
        { sheetId: sheet.id, address: address, formula: cell.raw },
      ]);
  }
  const snapshot = {
    inputs: inputs,
    formulas: formulas,
    signature: signature,
    volatile: [
      ...new Set(
        formulas.flatMap(([, formulaRecord]) =>
          findVolatileFunctionNames(formulaRecord.formula),
        ),
      ),
    ],
  };
  if (isDeepFrozenValue(sheet.cells)) {
    const snapshotsBySheetId = sheetInputCache.get(sheet.cells) ?? new Map();
    (snapshotsBySheetId.set(sheet.id, snapshot),
      sheetInputCache.set(sheet.cells, snapshotsBySheetId));
  }
  return snapshot;
}
function buildWorkbookCalculationDescriptor(workbook) {
  const calculationSettings = normalizeCalculationSettings(
    workbook.calculation ?? defaultCalculationSettings,
  );
  return JSON.stringify([
    workbook.id,
    calculationSettings,
    normalizeDatePolicy(workbook.datePolicy),
    workbook.definedNames,
    workbook.workbookLinks,
    workbook.extensions,
    getExtensionCacheEntries(workbook),
    workbook.sheets.map((sheet) => [
      sheet.id,
      sheet.name,
      sheet.rowCount,
      sheet.colCount,
      sheet.hiddenRows ?? [],
      sheet.hiddenCols ?? [],
      !!sheet.hidden,
      !!sheet.veryHidden,
      sheet.tables ?? [],
      sheet.merges ?? [],
      sheet.outlineGroups ?? [],
    ]),
  ]);
}
function buildWorkbookCalculationSignature(workbook) {
  const sheetSnapshots = workbook.sheets.map(buildSheetCalculationSnapshot);
  let cachedInputMap;
  const inputs = () =>
      (cachedInputMap ??= new Map(
        sheetSnapshots.flatMap((sheetSnapshot) => sheetSnapshot.inputs()),
      )),
    formulas = new Map(
      sheetSnapshots.flatMap((formulaSnapshot) => formulaSnapshot.formulas),
    ),
    calculationSettings = normalizeCalculationSettings(
      workbook.calculation ?? defaultCalculationSettings,
    ),
    workbookCalculationDescriptor =
      buildWorkbookCalculationDescriptor(workbook);
  let cachedInputSignature, cachedSignatureText;
  const inputSignature = () =>
      (cachedInputSignature ??= `[${JSON.stringify(workbook.revision)},${JSON.stringify(workbookCalculationDescriptor)},[${sheetSnapshots.map((inputSnapshot) => inputSnapshot.signature()).join(",")}]]`),
    signature = () =>
      (cachedSignatureText ??= `[${inputSignature()},${JSON.stringify(workbook.importedCalculationCaches ?? null)}]`);
  return (
    isDeepFrozenValue(workbook) || signature(),
    {
      inputs: inputs,
      formulas: formulas,
      settings: calculationSettings,
      metadata: workbookCalculationDescriptor,
      volatile: [
        ...new Set(
          sheetSnapshots
            .flatMap((volatileSnapshot) => volatileSnapshot.volatile)
            .concat(
              (workbook.definedNames ?? []).flatMap((definedName) =>
                findVolatileFunctionNames(definedName.refersTo),
              ),
            ),
        ),
      ].sort(),
      inputSignature: inputSignature,
      signature: signature,
    }
  );
}
function getCachedCalculationSnapshot(workbook) {
  if (!isDeepFrozenValue(workbook))
    return buildWorkbookCalculationSignature(workbook);
  let cachedSnapshot = calculationSnapshotCache.get(workbook);
  return (
    cachedSnapshot ||
      ((cachedSnapshot = buildWorkbookCalculationSignature(workbook)),
      calculationSnapshotCache.set(workbook, cachedSnapshot)),
    cachedSnapshot
  );
}
function isCalculationRecordCurrent(
  calculationRecord,
  workbook,
  cachedCalculationSnapshot = getCachedCalculationSnapshot(workbook),
) {
  return (
    !!calculationRecord &&
    (isDeepFrozenValue(workbook) ||
      calculationRecord.signature() === cachedCalculationSnapshot.signature())
  );
}
function MO(workbook, otherWorkbook) {
  if (
    workbook.revision !== otherWorkbook.revision ||
    buildWorkbookCalculationDescriptor(workbook) !==
      buildWorkbookCalculationDescriptor(otherWorkbook) ||
    (workbook.importedCalculationCaches !==
      otherWorkbook.importedCalculationCaches &&
      JSON.stringify(workbook.importedCalculationCaches) !==
        JSON.stringify(otherWorkbook.importedCalculationCaches))
  )
    return false;
  for (let sheetIndex = 0; sheetIndex < workbook.sheets.length; sheetIndex++) {
    const cells = workbook.sheets[sheetIndex].cells,
      otherCells = otherWorkbook.sheets[sheetIndex].cells;
    if (cells === otherCells) continue;
    const cellAddresses = Object.keys(cells);
    if (cellAddresses.length !== Object.keys(otherCells).length) return false;
    for (const address of cellAddresses) {
      const cell = cells[address],
        otherCell = otherCells[address];
      if (
        !otherCell ||
        cell.raw !== otherCell.raw ||
        !!cell.literal != !!otherCell.literal
      )
        return false;
    }
  }
  return true;
}
function hasCalculatorRecord(workbook) {
  return (
    calculatorRecords.has(workbook) || carriedCalculatorRecords.has(workbook)
  );
}
function carryCalculatorRecordForward(workbook, nextWorkbook) {
  if (
    workbook === nextWorkbook ||
    workbook.id !== nextWorkbook.id ||
    nextWorkbook.revision <= workbook.revision
  )
    return;
  const calculatorRecord = calculatorRecords.get(workbook);
  calculatorRecord &&
    isCalculationRecordCurrent(calculatorRecord, workbook) &&
    carriedCalculatorRecords.set(nextWorkbook, calculatorRecord);
}
function zie(previousWorkbook, workbook) {
  if (previousWorkbook === workbook) return;
  const previousRecord = calculatorRecords.get(previousWorkbook);
  previousRecord &&
  isCalculationRecordCurrent(previousRecord, previousWorkbook) &&
  MO(previousWorkbook, workbook)
    ? calculatorRecords.set(workbook, previousRecord)
    : carriedCalculatorRecords.has(previousWorkbook) &&
      MO(previousWorkbook, workbook) &&
      carriedCalculatorRecords.set(
        workbook,
        carriedCalculatorRecords.get(previousWorkbook),
      );
}
function withCalculationSettings(workbook, settings) {
  return { ...workbook, calculation: normalizeCalculationSettings(settings) };
}
function deepFreeze(value) {
  !value ||
    typeof value != "object" ||
    Object.isFrozen(value) ||
    (Object.values(value).forEach(deepFreeze), Object.freeze(value));
}
function getCurrentCalculatorRecord(activeWorkbook) {
  const calculatorRecord = calculatorRecords.get(activeWorkbook);
  return isCalculationRecordCurrent(calculatorRecord, activeWorkbook)
    ? {
        calculator: calculatorRecord.calculator,
        report: calculatorRecord.report,
      }
    : undefined;
}
const reportCellValueMaps = new WeakMap();
function buildReportCellValueMap(calculatorRecord) {
  if (!calculatorRecord) return new Map();
  const cachedValueMap = reportCellValueMaps.get(calculatorRecord.report);
  if (cachedValueMap) return cachedValueMap;
  const valueMap = new Map();
  for (const cell of calculatorRecord.report.cells) {
    const cellKey = `${cell.sheetId}:${cell.address}`,
      valueEntry = calculatorRecord.values.get(cellKey);
    valueEntry && valueEntry.value !== "#PENDING!"
      ? valueMap.set(cellKey, {
          revision: calculatorRecord.report.revision,
          result: { ...valueEntry },
        })
      : cell.lastKnown && valueMap.set(cellKey, cell.lastKnown);
  }
  return valueMap;
}
function getCalculationInputSignature(workbook) {
  return getCachedCalculationSnapshot(workbook).inputSignature();
}
const filterEvaluationLimits = Object.freeze({
    rows: 1e4,
    cells: 1e5,
    columns: 64,
    conditions: 128,
    values: 1e3,
    text: 1024,
    comparisonWork: 5e6,
  }),
  savedFilterViewHelpText =
    "View only. Column rules intersect (AND); advanced groups use OR of AND rules. Counts include blank records in the declared range. Manual/outline hiding is separate. Full scope includes every data row; visible scope excludes filtered and manually hidden rows. Applying a view leaves source addresses, formulas, identities and history unchanged. Native JSON, XLSX, print and AI selections use full source data. Active masks are session state; named criteria can be saved in native JSON and require a fresh preview before reuse.",
  normalizeComparableText = (text, isCaseSensitive = false) =>
    isCaseSensitive
      ? text.normalize("NFC")
      : text.normalize("NFC").toLocaleLowerCase("en-US"),
  comparableCellValueKey = (value, isCaseSensitive = false) =>
    isNullOrEmptyValue(value)
      ? "blank"
      : `${typeof value}:${typeof value == "string" ? normalizeComparableText(value, isCaseSensitive) : value}`,
  textFilterOperators = [
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
  ],
  isStorableCellValue = (value) =>
    value === null ||
    typeof value == "boolean" ||
    (typeof value == "number" && Number.isFinite(value)) ||
    (typeof value == "string" && value.length <= 32767);
function resolveDatePolicyContext(value, valueType, workbook) {
  if (valueType === "number") return typeof value == "number" ? value : null;
  if (valueType === "date") {
    const datePolicy = normalizeDatePolicy(workbook.datePolicy),
      serial =
        typeof value == "number"
          ? value
          : typeof value == "string"
            ? parseDateText(value, datePolicy)
            : undefined;
    return serial !== undefined &&
      serial >= 0 &&
      serial < dateSystemMaxSerial(datePolicy.dateSystem) + 1
      ? serial
      : null;
  }
  return typeof value == "string" ? value : null;
}
function parseWildcardPattern(text, pattern) {
  const tokens = [];
  for (let patternIndex = 0; patternIndex < pattern.length; patternIndex++) {
    const patternChar = pattern[patternIndex];
    patternChar === "~" && patternIndex + 1 < pattern.length
      ? tokens.push({ text: pattern[++patternIndex] })
      : tokens.push(
          patternChar === "*"
            ? { star: true }
            : patternChar === "?"
              ? { any: true }
              : { text: patternChar },
        );
  }
  let textIndex = 0,
    tokenIndex = 0,
    lastStarIndex = -1,
    starTextIndex = 0;
  for (; textIndex < text.length;)
    if (
      tokenIndex < tokens.length &&
      (tokens[tokenIndex].any || tokens[tokenIndex].text === text[textIndex])
    )
      (textIndex++, tokenIndex++);
    else if (tokenIndex < tokens.length && tokens[tokenIndex].star)
      ((lastStarIndex = tokenIndex++), (starTextIndex = textIndex));
    else if (lastStarIndex >= 0)
      ((tokenIndex = lastStarIndex + 1), (textIndex = ++starTextIndex));
    else return false;
  for (; tokenIndex < tokens.length && tokens[tokenIndex].star;) tokenIndex++;
  return tokenIndex === tokens.length;
}
function matchesTextFilter(value, filter, workbook, isCaseSensitive) {
  if (filter.operator === "blank") return isNullOrEmptyValue(value);
  if (filter.operator === "notBlank") return !isNullOrEmptyValue(value);
  if (filter.type === "value") {
    const isEqual =
      comparableCellValueKey(value, isCaseSensitive) ===
      comparableCellValueKey(filter.value, isCaseSensitive);
    return filter.operator === "equals" ? isEqual : !isEqual;
  }
  let datePolicyContext = resolveDatePolicyContext(
      value,
      filter.type,
      workbook,
    ),
    filterValue = resolveDatePolicyContext(filter.value, filter.type, workbook);
  if (datePolicyContext === null || filterValue === null) return false;
  switch (
    (typeof datePolicyContext == "string" &&
      typeof filterValue == "string" &&
      ((datePolicyContext = normalizeComparableText(
        datePolicyContext,
        isCaseSensitive,
      )),
      (filterValue = normalizeComparableText(filterValue, isCaseSensitive))),
    filter.operator)
  ) {
    case "equals":
      return datePolicyContext === filterValue;
    case "notEquals":
      return datePolicyContext !== filterValue;
    case "gt":
      return datePolicyContext > filterValue;
    case "gte":
      return datePolicyContext >= filterValue;
    case "lt":
      return datePolicyContext < filterValue;
    case "lte":
      return datePolicyContext <= filterValue;
    case "contains":
      return String(datePolicyContext).includes(String(filterValue));
    case "notContains":
      return !String(datePolicyContext).includes(String(filterValue));
    case "startsWith":
      return String(datePolicyContext).startsWith(String(filterValue));
    case "endsWith":
      return String(datePolicyContext).endsWith(String(filterValue));
    case "wildcard":
      return parseWildcardPattern(
        String(datePolicyContext),
        String(filterValue),
      );
  }
  return false;
}
function assertCalculatedBeforeFilter(
  getCalculatedValue,
  sheetId,
  address,
  workbook,
) {
  const value = getCalculatedValue(sheetId, address),
    cell = workbook?.sheets.find((sheet) => sheet.id === sheetId)?.cells[
      address
    ];
  if (
    cell &&
    cell.raw === value &&
    (cell.literal || typeof cell.raw != "string" || !cell.raw.startsWith("="))
  )
    return value;
  if (
    typeof value == "string" &&
    /^#(?:PENDING|LIMIT|TIMEOUT|CANCELLED|BUSY|BLOCKED|CALC_LIMIT)!?$/i.test(
      value,
    )
  )
    throw Error(
      `Calculation is incomplete at ${address}. Recalculate before filtering; unchecked rows are never reported as an empty population.`,
    );
  return value;
}
const savedFilterViewLimits = Object.freeze({
  views: 20,
  name: 80,
  characters: 512e3,
});
function validateSavedFilterViewMetadata(metadata, allowedKeys) {
  if (
    !metadata ||
    typeof metadata != "object" ||
    Array.isArray(metadata) ||
    ![Object.prototype, null].includes(Object.getPrototypeOf(metadata)) ||
    Object.keys(metadata).some((key) => !allowedKeys.includes(key))
  )
    throw Error("Invalid saved filter view metadata.");
}
function isValidFilterViewText(text, maxLength) {
  return (
    typeof text == "string" &&
    !!text.trim() &&
    text.length <= maxLength &&
    !/[\u0000-\u001f\u007f]/.test(text)
  );
}
function normalizeNameForComparison(savedFilterViewName) {
  return savedFilterViewName.normalize("NFC").trim().toLocaleLowerCase("en-US");
}
function normalizeSavedFilterViewName(viewName) {
  if (!isValidFilterViewText(viewName, savedFilterViewLimits.name))
    throw Error("Use a nonempty filter view name of up to 80 characters.");
  return viewName.normalize("NFC").trim();
}
function assertSavedFilterViewsEditable(workbook, sheetId) {
  if (
    workbook.structureProtection?.enabled ||
    workbook.sheets.find((sheet) => sheet.id === sheetId)?.protection?.enabled
  )
    throw Error(
      "Unprotect the workbook and source worksheet before changing saved filter views.",
    );
}
function findSavedFilterView(workbook, viewId) {
  const view = workbook.savedFilterViews?.find(
    (savedFilterView) => savedFilterView.id === viewId,
  );
  if (!view) throw Error("This saved filter view no longer exists.");
  return view;
}
const consolidationFormatVersion = "ledgerglass-consolidation-v1",
  consolidationLimits = Object.freeze({
    sources: 8,
    categories: 3,
    contributions: 2e3,
    outputRows: 4e3,
    outputCells: 4e4,
    groups: 1e3,
    metadataCharacters: 8e6,
  }),
  boundedLabelSchema = string()
    .min(1)
    .max(128)
    .refine((text) => !!text.trim() && !/[\u0000-\u001f\u007f]/.test(text)),
  shortLabelSchema = string().min(1).max(100),
  V5 = string().regex(/^[a-f0-9]{64}$/),
  scalarValueSchema = union([
    string().min(1).max(32767),
    number().finite(),
    boolean(),
  ]),
  categoryValueSchema = discriminatedUnion("kind", [
    strictObject({ kind: literal("constant"), value: boundedLabelSchema }),
    strictObject({ kind: literal("field"), field: boundedLabelSchema }),
  ]),
  consolidationSourceSchema = strictObject({
    id: shortLabelSchema,
    workbookId: shortLabelSchema,
    sheetId: shortLabelSchema,
    range: string().min(1).max(50),
    tableId: shortLabelSchema.optional(),
    categories: array(boundedLabelSchema).min(1).max(3),
    amountField: boundedLabelSchema,
    period: categoryValueSchema,
    currency: categoryValueSchema,
    unit: categoryValueSchema,
    duplicatePolicy: _enum(["reject", "sum"]),
  }),
  consolidationDefinitionSchema = strictObject({
    engine: literal(consolidationFormatVersion),
    sources: array(consolidationSourceSchema).min(1).max(8),
    categoryLabels: array(boundedLabelSchema).min(1).max(3),
    period: boundedLabelSchema,
    currency: boundedLabelSchema,
    unit: boundedLabelSchema,
    linkedAmounts: boolean(),
    collapseDetails: boolean(),
  }).superRefine((definition, context) => {
    const outputHeaders = [
      "Row type",
      ...definition.categoryLabels,
      "Period",
      "Currency",
      "Unit",
      "Amount",
      "Source",
      "Link state",
    ];
    (new Set(outputHeaders.map((header) => header.toLowerCase())).size !==
      outputHeaders.length ||
      new Set(definition.sources.map((source) => source.id)).size !==
        definition.sources.length ||
      definition.sources.some(
        (source) =>
          source.categories.length !== definition.categoryLabels.length ||
          new Set(source.categories).size !== source.categories.length,
      )) &&
      context.addIssue({
        code: "custom",
        message:
          "Use unique source IDs, distinct category fields, and distinct output headers with matching category counts.",
      });
  }),
  consolidationSourceSnapshotSchema = strictObject({
    sourceId: shortLabelSchema,
    workbookId: shortLabelSchema,
    workbookName: string().min(1).max(200),
    sheetId: shortLabelSchema,
    sheetName: string().min(1).max(31),
    row: number().int().min(2).max(1048576),
    address: string().regex(/^[A-Z]{1,3}[1-9]\d{0,6}$/),
    locator: string().max(300),
    categories: array(scalarValueSchema).min(1).max(3),
    amount: number().finite(),
    period: boundedLabelSchema,
    currency: boundedLabelSchema,
    unit: boundedLabelSchema,
    link: _enum(["live", "snapshot", "approved snapshot"]),
  }),
  consolidationIssueSchema = strictObject({
    sourceId: shortLabelSchema,
    row: number().int().positive().optional(),
    code: string().max(40),
    severity: _enum(["error", "warning"]),
    message: string().max(4e3),
  }),
  consolidationSourceResultSchema = strictObject({
    sourceId: shortLabelSchema,
    workbookId: shortLabelSchema,
    workbookName: string().max(200),
    sheetId: shortLabelSchema,
    sheetName: string().max(31),
    range: string().max(50),
    fingerprint: V5,
    records: number().int().min(0).max(500),
    acceptedRecords: number().int().min(0).max(500),
    amount: number().finite(),
    periods: array(boundedLabelSchema).max(500),
    currencies: array(boundedLabelSchema).max(500),
    units: array(boundedLabelSchema).max(500),
  }),
  consolidationSubtotalSchema = strictObject({
    categories: array(scalarValueSchema).min(1).max(3),
    level: number().int().min(1).max(3),
    count: number().int().min(1).max(2e3),
    amount: number().finite(),
  }),
  consolidationTotalsSchema = strictObject({
    sources: array(consolidationSourceResultSchema).max(8),
    contributions: array(consolidationSourceSnapshotSchema).max(2e3),
    subtotals: array(consolidationSubtotalSchema).max(1e3),
    exceptions: array(consolidationIssueSchema).max(8e3),
    grandTotal: number().finite().nullable(),
    sourceTotal: number().finite().nullable(),
    difference: number().finite().nullable(),
    complete: boolean(),
  }),
  consolidationEvidenceSchema = strictObject({
    engine: literal(consolidationFormatVersion),
    definition: consolidationDefinitionSchema,
    review: consolidationTotalsSchema,
    origin: strictObject({
      workbookId: shortLabelSchema,
      revision: number().int().nonnegative(),
      at: string().datetime(),
    }),
    evidenceHash: V5,
    outputRange: string().max(50),
    outputHash: V5,
  });
function parseConsolidationDefinition(definition) {
  const parsed = consolidationDefinitionSchema.safeParse(definition);
  if (!parsed.success)
    throw Error(`Invalid consolidation: ${parsed.error.issues[0]?.message}`);
  return parsed.data;
}
const serializeTypedKeyValues = (entries) =>
  JSON.stringify(entries.map((item) => [typeof item, item]));
function sumConsolidationAmounts(amounts) {
  let total = 0,
    compensation = 0;
  for (const amount of amounts) {
    const adjustedAmount = amount - compensation,
      nextTotal = total + adjustedAmount;
    if (
      ((compensation = nextTotal - total - adjustedAmount),
      (total = nextTotal),
      !Number.isFinite(total) || !Number.isFinite(compensation))
    )
      throw Error("Consolidation total exceeds finite numeric limits.");
  }
  return Object.is(total, -0) ? 0 : total;
}
function buildConsolidationSubtotals(contributions) {
  const groupsByKey = new Map();
  for (const contribution of contributions)
    for (
      let prefixLength = 1;
      prefixLength <= contribution.categories.length;
      prefixLength++
    ) {
      const groupKey = serializeTypedKeyValues(
          contribution.categories.slice(0, prefixLength),
        ),
        groupContributions = groupsByKey.get(groupKey) ?? [];
      (groupContributions.push(contribution),
        groupsByKey.set(groupKey, groupContributions));
    }
  if (groupsByKey.size > consolidationLimits.groups)
    throw Error(
      "Consolidation exceeds 1,000 subtotal groups. Narrow the sources; no rows were truncated.",
    );
  return [...groupsByKey].map(([serializedKey, groupMembers]) => {
    const level = JSON.parse(serializedKey).length;
    return {
      categories: groupMembers[0].categories.slice(0, level),
      level: level,
      count: groupMembers.length,
      amount: sumConsolidationAmounts(
        groupMembers.map((entry) => entry.amount),
      ),
    };
  });
}
function verifyConsolidationEvidence(workbook, report) {
  let totalMetadataCharacters = 0;
  for (const sheet of report.sheets) {
    const sheetConsolidation = workbook.sheets?.find(
      (workbookSheet) => workbookSheet.id === sheet.id,
    )?.consolidation;
    if (sheetConsolidation === undefined) continue;
    if (
      ((totalMetadataCharacters += JSON.stringify(sheetConsolidation).length),
      totalMetadataCharacters > consolidationLimits.metadataCharacters)
    )
      throw Error(
        "Consolidation evidence exceeds 8,000,000 characters. Retain native backups before removing obsolete reports.",
      );
    const parsedEvidence =
      consolidationEvidenceSchema.safeParse(sheetConsolidation);
    if (!parsedEvidence.success)
      throw Error(
        `Invalid consolidation evidence: ${parsedEvidence.error.issues[0]?.message}`,
      );
    const data = parsedEvidence.data,
      review = data.review,
      sourceIds = data.definition.sources.map((source) => source.id);
    if (
      !review.complete ||
      review.exceptions.some((exception) => exception.severity === "error") ||
      !review.contributions.length ||
      review.sources.length !== sourceIds.length ||
      new Set(review.sources.map((source) => source.sourceId)).size !==
        sourceIds.length ||
      review.sources.some((source) => !sourceIds.includes(source.sourceId)) ||
      review.contributions.some(
        (contribution) =>
          !sourceIds.includes(contribution.sourceId) ||
          contribution.categories.length !==
            data.definition.categoryLabels.length ||
          contribution.period !== data.definition.period ||
          contribution.currency !== data.definition.currency ||
          contribution.unit !== data.definition.unit,
      ) ||
      review.grandTotal !==
        sumConsolidationAmounts(
          review.contributions.map((contribution) => contribution.amount),
        ) ||
      review.sourceTotal !==
        sumConsolidationAmounts(
          review.sources.map((source) => source.amount),
        ) ||
      review.difference !== review.grandTotal - review.sourceTotal ||
      fingerprintCanonicalValue(review.subtotals) !==
        fingerprintCanonicalValue(
          buildConsolidationSubtotals(review.contributions),
        ) ||
      data.evidenceHash !==
        fingerprintCanonicalValue({
          definition: data.definition,
          review: review,
          origin: data.origin,
        })
    )
      throw Error(
        "Consolidation evidence does not reconcile to its sources, categories or retained fingerprint.",
      );
    for (const source of review.sources) {
      const sourceContributions = review.contributions.filter(
        (contribution) => contribution.sourceId === source.sourceId,
      );
      if (
        source.records !== sourceContributions.length ||
        source.acceptedRecords !== sourceContributions.length ||
        source.amount !==
          sumConsolidationAmounts(
            sourceContributions.map((entry) => entry.amount),
          ) ||
        sourceContributions.some(
          (sourceContribution) =>
            sourceContribution.workbookId !== source.workbookId ||
            sourceContribution.sheetId !== source.sheetId,
        )
      )
        throw Error("Consolidation source evidence does not reconcile.");
    }
    sheet.consolidation = data;
  }
}
const externalSourceFormatVersion = "ledgerglass-sources-v1",
  externalSourceKinds = [
    "database",
    "odbc",
    "odata",
    "api",
    "erp",
    "file-folder",
  ],
  externalSourceLimits = Object.freeze({
    sources: 20,
    columns: 50,
    rows: 5e3,
    cells: 1e5,
    responseBytes: 8 * 1024 * 1024,
    timeoutMs: 3e4,
    metadataCharacters: 2e5,
  }),
  sourceIdSchema = string().regex(/^[A-Za-z0-9][A-Za-z0-9._-]{0,99}$/),
  columnNameSchema = string()
    .min(1)
    .max(128)
    .refine(
      (columnName) =>
        columnName === columnName.trim() &&
        !/[\u0000-\u001f\u007f]/.test(columnName),
    ),
  isoDateTimeSchema = string().datetime({ offset: true }),
  sourceHashSchema = string().regex(/^[a-f0-9]{64}$/),
  sourceColumnSchema = strictObject({
    name: columnNameSchema,
    type: _enum(["text", "number", "boolean", "date"]),
    nullable: boolean(),
  }),
  sourceTableSchema = strictObject({
    id: sourceIdSchema,
    label: columnNameSchema,
    kind: _enum(externalSourceKinds),
    schemaVersion: sourceIdSchema,
    columns: array(sourceColumnSchema).min(1).max(externalSourceLimits.columns),
    controlFields: array(columnNameSchema).max(externalSourceLimits.columns),
  }).superRefine((sourceTable, context) => {
    (new Set(sourceTable.columns.map((column) => column.name.toLowerCase()))
      .size !== sourceTable.columns.length ||
      new Set(sourceTable.controlFields).size !==
        sourceTable.controlFields.length ||
      sourceTable.controlFields.some(
        (controlField) =>
          !sourceTable.columns.some(
            (column) =>
              column.name === controlField && column.type === "number",
          ),
      )) &&
      context.addIssue({
        code: "custom",
        message: "Use unique columns and declared numeric control fields.",
      });
  }),
  sourceApprovalSchema = strictObject({
    source: sourceTableSchema,
    expiresAt: isoDateTimeSchema,
  }),
  sourceCellValueSchema = union([
    string().max(32767),
    number().finite(),
    boolean(),
    _null(),
  ]),
  sourceResultPayloadSchema = strictObject({
    protocol: literal(externalSourceFormatVersion),
    source: sourceTableSchema,
    sourceTimestamp: isoDateTimeSchema,
    version: sourceIdSchema,
    rows: array(
      array(sourceCellValueSchema).max(externalSourceLimits.columns),
    ).max(externalSourceLimits.rows),
    complete: literal(true),
  }),
  externalSourceErrorCodes = [
    "unavailable",
    "refused",
    "expired",
    "schema",
    "limit",
    "timeout",
    "cancelled",
    "invalid",
    "stale",
  ],
  externalSourceErrorMessages = {
    unavailable:
      "The source is unavailable. The last successful data is retained. Check the host connection, then retry.",
    refused:
      "This source is not authorized by the host. No source data was replaced.",
    expired:
      "Source access expired. Renew access in the host application and preview again.",
    schema:
      "The source schema changed. The working population is retained. Have the host approve the new schema and import it into a new worksheet.",
    limit:
      "The source exceeds the published row, column, cell or response limit. No rows were truncated or published.",
    timeout:
      "Source refresh timed out. The last successful data is retained; retry when the host is available.",
    cancelled:
      "Source refresh cancelled. The last successful data is retained.",
    invalid:
      "The source returned an invalid or incomplete result. The last successful data is retained.",
    stale:
      "The workbook, preview or source approval changed. Preview again before applying a refresh.",
  };
class ExternalSourceError extends Error {
  constructor(errorCode) {
    (super(externalSourceErrorMessages[errorCode]),
      (this.code = errorCode),
      (this.name = "ExternalSourceError"));
  }
}
function isExternalSourceError(error) {
  return error instanceof ExternalSourceError
    ? error
    : error instanceof Error &&
        error.name === "ExternalSourceError" &&
        externalSourceErrorCodes.includes(error.code)
      ? new ExternalSourceError(error.code)
      : new ExternalSourceError("unavailable");
}
function parseActiveSourceApproval(approval, currentTimeMs = Date.now()) {
  const parseResult = sourceApprovalSchema.safeParse(approval);
  if (!parseResult.success) throw new ExternalSourceError("invalid");
  if (Date.parse(parseResult.data.expiresAt) <= currentTimeMs)
    throw new ExternalSourceError("expired");
  return parseResult.data;
}
function isSameContentHash(source, otherSource) {
  return (
    fingerprintCanonicalValue(source) === fingerprintCanonicalValue(otherSource)
  );
}
function validateExternalSourcePayload(snapshot, source) {
  if (snapshot && typeof snapshot == "object") {
    const rows = snapshot.rows;
    if (
      Array.isArray(rows) &&
      (rows.length > externalSourceLimits.rows ||
        rows.some(
          (row) =>
            Array.isArray(row) && row.length > externalSourceLimits.columns,
        ))
    )
      throw new ExternalSourceError("limit");
  }
  const parseResult = sourceResultPayloadSchema.safeParse(snapshot);
  if (!parseResult.success) throw new ExternalSourceError("invalid");
  const data = parseResult.data;
  if (source && !isSameContentHash(data.source, source))
    throw new ExternalSourceError("schema");
  if (
    (data.rows.length + 1) * data.source.columns.length >
      externalSourceLimits.cells ||
    new TextEncoder().encode(JSON.stringify(data)).length >
      externalSourceLimits.responseBytes
  )
    throw new ExternalSourceError("limit");
  for (const row of data.rows) {
    if (row.length !== data.source.columns.length)
      throw new ExternalSourceError("schema");
    row.forEach((cellValue, columnIndex) => {
      const column = data.source.columns[columnIndex];
      if (
        cellValue === null
          ? !column.nullable
          : column.type === "date"
            ? typeof cellValue != "string" ||
              !/^\d{4}-\d{2}-\d{2}$/.test(cellValue) ||
              !Number.isFinite(Date.parse(cellValue)) ||
              new Date(cellValue).toISOString().slice(0, 10) !== cellValue
            : typeof cellValue !==
              { text: "string", number: "number", boolean: "boolean" }[
                column.type
              ]
      )
        throw new ExternalSourceError("schema");
    });
  }
  return data;
}
function computeSourceControlTotals(snapshot) {
  return snapshot.source.controlFields.map((controlField) => {
    const controlColumnIndex = snapshot.source.columns.findIndex(
      (column) => column.name === controlField,
    );
    let total = 0,
      compensation = 0,
      numeric = 0,
      blanks = 0;
    for (const row of snapshot.rows) {
      if (row[controlColumnIndex] === null) {
        blanks++;
        continue;
      }
      const adjustedValue = row[controlColumnIndex] - compensation,
        nextTotal = total + adjustedValue;
      if (
        ((compensation = nextTotal - total - adjustedValue),
        (total = nextTotal),
        numeric++,
        !Number.isFinite(total) || !Number.isFinite(compensation))
      )
        throw new ExternalSourceError("limit");
    }
    return {
      field: controlField,
      sum: Object.is(total, -0) ? 0 : total,
      numeric: numeric,
      blanks: blanks,
    };
  });
}
const sourceControlFieldSchema = strictObject({
    field: columnNameSchema,
    sum: number().finite(),
    numeric: number().int().min(0).max(5e3),
    blanks: number().int().min(0).max(5e3),
  }),
  sourceRefreshStateSchema = strictObject({
    protocol: literal(externalSourceFormatVersion),
    source: sourceTableSchema,
    lastSuccess: strictObject({
      at: isoDateTimeSchema,
      sourceTimestamp: isoDateTimeSchema,
      version: sourceIdSchema,
      rows: number().int().min(0).max(5e3),
      controls: array(sourceControlFieldSchema).max(50),
      dataHash: sourceHashSchema,
    }),
    lastAttempt: strictObject({
      at: isoDateTimeSchema,
      outcome: _enum(["succeeded", "failed", "cancelled"]),
      code: _enum(externalSourceErrorCodes).optional(),
    }),
    outputHash: sourceHashSchema,
  });
function buildExternalSourceRefreshState(workbook, report) {
  let totalMetadataCharacters = 0,
    sourceCount = 0;
  for (const sheet of report.sheets) {
    const externalConnection = workbook.sheets.find(
      (workbookSheet) => workbookSheet.id === sheet.id,
    )?.externalConnection;
    if (externalConnection === undefined) continue;
    if (
      ++sourceCount > externalSourceLimits.sources ||
      (totalMetadataCharacters += JSON.stringify(externalConnection).length) >
        externalSourceLimits.metadataCharacters
    )
      throw new ExternalSourceError("limit");
    const parseResult = sourceRefreshStateSchema.safeParse(externalConnection);
    if (!parseResult.success) throw new ExternalSourceError("invalid");
    const data = parseResult.data;
    if (
      data.lastSuccess.controls.length !== data.source.controlFields.length ||
      data.lastSuccess.controls.some(
        (control, controlIndex) =>
          control.field !== data.source.controlFields[controlIndex] ||
          control.numeric + control.blanks !== data.lastSuccess.rows,
      ) ||
      (data.lastSuccess.rows + 1) * data.source.columns.length >
        externalSourceLimits.cells ||
      Date.parse(data.lastAttempt.at) < Date.parse(data.lastSuccess.at) ||
      (data.lastAttempt.outcome === "succeeded") == !!data.lastAttempt.code
    )
      throw new ExternalSourceError("invalid");
    sheet.externalConnection = data;
  }
}

export {
  assertCalculatedBeforeFilter,
  hasCalculatorRecord,
  deepFreeze,
  isSameContentHash,
  formatNumberWithPattern,
  normalizeNameForComparison,
  consolidationFormatVersion,
  filterEvaluationLimits,
  serializeTypedKeyValues,
  defaultCalculationSettings,
  makeFormulaError,
  createWorkbookFormulaEngine,
  savedFilterViewHelpText,
  getCurrentCalculatorRecord,
  matchesTextFilter,
  importFormulaCacheEvidence,
  isExternalSourceError,
  normalizeSavedFilterViewName,
  normalizeCalculationSettings,
  externalSourceFormatVersion,
  sumConsolidationAmounts,
  reportCellValueMaps,
  assertSavedFilterViewsEditable,
  ExternalSourceError,
  isValidFilterViewText,
  carryCalculatorRecordForward,
  applyDefinedNameOperation,
  resolveDatePolicyContext,
  withCalculationSettings,
  getCachedCalculationSnapshot,
  textFilterOperators,
  normalizeComparableText,
  _O,
  findSavedFilterView,
  validateExternalSourcePayload,
  buildExternalSourceRefreshState,
  carriedCalculatorRecords,
  comparableCellValueKey,
  externalSourceLimits,
  buildConsolidationSubtotals,
  parseActiveSourceApproval,
  ik,
  j5,
  isStorableCellValue,
  consolidationLimits,
  getCalculationInputSignature,
  parseConsolidationDefinition,
  pendingFormulaValue,
  isCalculationRecordCurrent,
  assertNamePreviewCurrent,
  buildReportCellValueMap,
  calculatorRecords,
  verifyConsolidationEvidence,
  validateSavedFilterViewMetadata,
  computeSourceControlTotals,
  savedFilterViewLimits,
  zie,
};

/* Also used by the workers. */
export {
  MO,
  boundedLabelSchema,
  buildSheetCalculationSnapshot,
  buildWorkbookCalculationDescriptor,
  buildWorkbookCalculationSignature,
  categoryValueSchema,
  columnNameSchema,
  consolidationDefinitionSchema,
  consolidationEvidenceSchema,
  consolidationIssueSchema,
  consolidationSourceResultSchema,
  consolidationSourceSchema,
  consolidationSourceSnapshotSchema,
  consolidationSubtotalSchema,
  consolidationTotalsSchema,
  dynamicFunctionPattern,
  externalSourceErrorCodes,
  externalSourceKinds,
  isoDateTimeSchema,
  matchCriteriaValue,
  scalarValueSchema,
  sourceApprovalSchema,
  sourceColumnSchema,
  sourceControlFieldSchema,
  sourceIdSchema,
  sourceRefreshStateSchema,
  sourceResultPayloadSchema,
  sourceTableSchema,
  throwInvalidCalculationSettings,
};

/* Also used by the workers. */
export { externalSourceErrorMessages };
