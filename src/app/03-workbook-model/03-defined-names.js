import {
  coerceValueToText,
  flattenNestedValues,
} from "../02-text-and-financial-functions.js";
import {
  isNonArrayObject,
  failWithMessage,
  throwWorkbookError,
} from "./01-cell-styles-and-themes.js";
import {
  walkFormulaNodes,
  tokenizeFormulaText,
  normalizeBindingName,
  validateLetLambdaCall,
  validateBindingName,
  normalizeFunctionName,
  dropBlankTextTokens,
  sha256HexDigest,
} from "../01-schemas-date-functions-and-validation.js";
import {
  toMatrixArgument,
  valuesEqualIgnoringCase,
  compareSortKeyWithDirection,
  compareSortValues,
  notModernFunctionSentinel,
  modernFunctionNames,
  matrixResultFunctionNames,
} from "./02-formula-function-evaluators.js";

function evaluateDefinedNameFunction(functionName, argRefs, evaluationContext) {
  if (!modernFunctionNames.includes(functionName))
    return notModernFunctionSentinel;
  const raiseIssue = (errorCode = "#VALUE!") => {
      throw evaluationContext.issue(errorCode);
    },
    coerceNumberArg = (argValue, defaultValue) => {
      if (evaluationContext.isIssue(argValue)) throw argValue;
      if (argValue === undefined && defaultValue !== undefined)
        return defaultValue;
      if (
        Array.isArray(argValue) ||
        (typeof argValue == "string" && !argValue.trim())
      )
        return raiseIssue();
      const numericValue = Number(argValue);
      return Number.isFinite(numericValue) ? numericValue : raiseIssue();
    },
    coerceIntegerArg = (rawArg, defaultArg) =>
      Math.trunc(coerceNumberArg(rawArg, defaultArg)),
    validateArrayShape = (shapeRows, shapeColumns) => {
      ((shapeRows < 1 || shapeColumns < 1) && raiseIssue("#CALC!"),
        shapeRows * shapeColumns > 1e4 && raiseIssue("#LIMIT!"));
    },
    coerceBooleanArg = (flagValue) => {
      if (evaluationContext.isIssue(flagValue)) throw flagValue;
      return (
        Array.isArray(flagValue) && raiseIssue(),
        typeof flagValue == "string" &&
          !/^(true|false)$/i.test(flagValue) &&
          raiseIssue(),
        typeof flagValue == "string"
          ? flagValue.toLowerCase() === "true"
          : !!flagValue
      );
    };
  if (functionName === "IFS") {
    (argRefs.length < 2 || argRefs.length % 2) && raiseIssue();
    for (
      let conditionIndex = 0;
      conditionIndex < argRefs.length;
      conditionIndex += 2
    ) {
      const conditionValue = evaluationContext.retrieve(
        argRefs[conditionIndex],
      );
      if (
        (Array.isArray(conditionValue) && raiseIssue(),
        coerceBooleanArg(conditionValue))
      )
        return evaluationContext.retrieve(argRefs[conditionIndex + 1]) ?? 0;
    }
    return raiseIssue("#N/A");
  }
  if (functionName === "SWITCH") {
    argRefs.length < 3 && raiseIssue();
    const switchValue = evaluationContext.retrieve(argRefs[0]);
    if (evaluationContext.isIssue(switchValue)) return switchValue;
    Array.isArray(switchValue) && raiseIssue();
    for (let pairIndex = 1; pairIndex + 1 < argRefs.length; pairIndex += 2) {
      const caseValue = evaluationContext.retrieve(argRefs[pairIndex]);
      if (evaluationContext.isIssue(caseValue)) return caseValue;
      if (
        (Array.isArray(caseValue) && raiseIssue(),
        valuesEqualIgnoringCase(switchValue, caseValue))
      )
        return evaluationContext.retrieve(argRefs[pairIndex + 1]) ?? 0;
    }
    return argRefs.length % 2 === 0
      ? (evaluationContext.retrieve(argRefs.at(-1)) ?? 0)
      : raiseIssue("#N/A");
  }
  const retrievedArgs = argRefs.map(evaluationContext.retrieve),
    arityBounds = {
      SEQUENCE: [1, 4],
      FILTER: [2, 3],
      TRANSPOSE: [1, 1],
      SORT: [1, 4],
      SORTBY: [2, 255],
      UNIQUE: [1, 3],
      TAKE: [2, 3],
      DROP: [2, 3],
      HSTACK: [1, 254],
      VSTACK: [1, 254],
    }[functionName];
  if (
    (arityBounds &&
      (argRefs.length < arityBounds[0] || argRefs.length > arityBounds[1]) &&
      raiseIssue(),
    arityBounds &&
      functionName !== "SEQUENCE" &&
      retrievedArgs[0] === undefined &&
      raiseIssue(),
    arityBounds &&
      !["HSTACK", "VSTACK"].includes(functionName) &&
      evaluationContext.isIssue(retrievedArgs[0]))
  )
    throw retrievedArgs[0];
  if (
    !matrixResultFunctionNames.includes(functionName) &&
    functionName !== "MAXIFS" &&
    functionName !== "MINIFS"
  ) {
    const firstArgError = retrievedArgs
      .flatMap(flattenNestedValues)
      .find(evaluationContext.isIssue);
    if (firstArgError) return firstArgError;
  }
  if (
    [
      "STDEV.P",
      "VAR.P",
      "PERCENTILE.INC",
      "PERCENTILE.EXC",
      "QUARTILE.INC",
      "QUARTILE.EXC",
      "RANK.EQ",
      "RANK.AVG",
      "MODE.SNGL",
    ].includes(functionName)
  ) {
    const isRankFunction = functionName.startsWith("RANK"),
      sortedNumbers = (
        functionName === "STDEV.P" ||
        functionName === "VAR.P" ||
        functionName === "MODE.SNGL"
          ? retrievedArgs.flatMap(flattenNestedValues)
          : flattenNestedValues(retrievedArgs[isRankFunction ? 1 : 0])
      )
        .filter((entry) => typeof entry == "number")
        .sort((leftValue, rightValue) => leftValue - rightValue);
    if ((sortedNumbers.length || raiseIssue("#NUM!"), isRankFunction)) {
      const rankValue = coerceNumberArg(retrievedArgs[0]),
        rankOrder = coerceNumberArg(retrievedArgs[2], 0),
        firstIndex = sortedNumbers.indexOf(rankValue),
        lastIndex = sortedNumbers.lastIndexOf(rankValue);
      return (
        firstIndex < 0 && raiseIssue("#N/A"),
        rankOrder
          ? 1 +
            (functionName === "RANK.AVG"
              ? (firstIndex + lastIndex) / 2
              : firstIndex)
          : sortedNumbers.length -
            (functionName === "RANK.AVG"
              ? (firstIndex + lastIndex) / 2
              : lastIndex)
      );
    }
    if (functionName === "MODE.SNGL") {
      const frequencyByValue = new Map();
      sortedNumbers.forEach((sampleValue) =>
        frequencyByValue.set(
          sampleValue,
          (frequencyByValue.get(sampleValue) ?? 0) + 1,
        ),
      );
      const frequencyEntries = [...frequencyByValue].sort(
        (leftEntry, rightEntry) =>
          rightEntry[1] - leftEntry[1] || leftEntry[0] - rightEntry[0],
      );
      return frequencyEntries[0][1] > 1
        ? frequencyEntries[0][0]
        : raiseIssue("#N/A");
    }
    if (functionName === "STDEV.P" || functionName === "VAR.P") {
      const populationMean =
          sortedNumbers.reduce((total, numberValue) => total + numberValue, 0) /
          sortedNumbers.length,
        populationVariance =
          sortedNumbers.reduce(
            (varianceTotal, valueItem) =>
              varianceTotal + (valueItem - populationMean) ** 2,
            0,
          ) / sortedNumbers.length;
      return functionName === "STDEV.P"
        ? Math.sqrt(populationVariance)
        : populationVariance;
    }
    const percentileFraction =
        coerceNumberArg(retrievedArgs[1]) /
        (functionName.startsWith("QUARTILE") ? 4 : 1),
      isExclusive = functionName.endsWith("EXC"),
      fractionalIndex = isExclusive
        ? percentileFraction * (sortedNumbers.length + 1) - 1
        : percentileFraction * (sortedNumbers.length - 1);
    (percentileFraction < 0 ||
      percentileFraction > 1 ||
      fractionalIndex < 0 ||
      fractionalIndex > sortedNumbers.length - 1) &&
      raiseIssue("#NUM!");
    const lowerIndex = Math.floor(fractionalIndex);
    return (
      sortedNumbers[lowerIndex] +
      (fractionalIndex - lowerIndex) *
        ((sortedNumbers[lowerIndex + 1] ?? sortedNumbers[lowerIndex]) -
          sortedNumbers[lowerIndex])
    );
  }
  if (functionName === "MDETERM" || functionName === "MINVERSE") {
    argRefs.length !== 1 && raiseIssue();
    const matrixArgument = toMatrixArgument(retrievedArgs[0]),
      matrixSize = matrixArgument.length;
    (matrixSize > 32 && raiseIssue("#LIMIT!"),
      (!matrixSize ||
        matrixArgument.some(
          (matrixRowToCheck) => matrixRowToCheck.length !== matrixSize,
        )) &&
        raiseIssue());
    for (const cellValue of matrixArgument.flat()) {
      if (evaluationContext.isIssue(cellValue)) return cellValue;
      (typeof cellValue != "number" && raiseIssue(),
        Number.isFinite(cellValue) || raiseIssue("#NUM!"));
    }
    const augmentedMatrix = matrixArgument.map(
      (sourceRow, augmentedRowIndex) => [
        ...sourceRow.map((cell) => cell),
        ...Array.from({ length: matrixSize }, (y, identityColumn) =>
          augmentedRowIndex === identityColumn ? 1 : 0,
        ),
      ],
    );
    let determinant = 1;
    for (let pivotIndex = 0; pivotIndex < matrixSize; pivotIndex++) {
      let pivotRow = pivotIndex;
      for (
        let candidateRow = pivotIndex + 1;
        candidateRow < matrixSize;
        candidateRow++
      )
        Math.abs(augmentedMatrix[candidateRow][pivotIndex]) >
          Math.abs(augmentedMatrix[pivotRow][pivotIndex]) &&
          (pivotRow = candidateRow);
      if (augmentedMatrix[pivotRow][pivotIndex] === 0)
        return functionName === "MDETERM" ? 0 : raiseIssue("#NUM!");
      pivotRow !== pivotIndex &&
        (([augmentedMatrix[pivotIndex], augmentedMatrix[pivotRow]] = [
          augmentedMatrix[pivotRow],
          augmentedMatrix[pivotIndex],
        ]),
        (determinant *= -1));
      const pivotValue = augmentedMatrix[pivotIndex][pivotIndex];
      ((determinant *= pivotValue),
        (augmentedMatrix[pivotIndex] = augmentedMatrix[pivotIndex].map(
          (rowCell) => rowCell / pivotValue,
        )));
      for (
        let eliminationRow = 0;
        eliminationRow < matrixSize;
        eliminationRow++
      )
        if (eliminationRow !== pivotIndex) {
          const eliminationFactor = augmentedMatrix[eliminationRow][pivotIndex];
          augmentedMatrix[eliminationRow] = augmentedMatrix[eliminationRow].map(
            (targetCell, elementIndex) =>
              targetCell -
              eliminationFactor * augmentedMatrix[pivotIndex][elementIndex],
          );
        }
    }
    const inverseRows = augmentedMatrix.map((matrixRow) =>
      matrixRow.slice(matrixSize),
    );
    return functionName === "MDETERM"
      ? Number.isFinite(determinant)
        ? determinant
        : raiseIssue("#NUM!")
      : (inverseRows.some((inverseRow) =>
          inverseRow.some((inverseCell) => !Number.isFinite(inverseCell)),
        ) && raiseIssue("#NUM!"),
        inverseRows);
  }
  if (functionName === "CONCAT" || functionName === "TEXTJOIN") {
    functionName === "TEXTJOIN" && retrievedArgs.length < 3 && raiseIssue();
    const delimiterText =
        functionName === "TEXTJOIN" ? coerceValueToText(retrievedArgs[0]) : "",
      shouldIgnoreEmpty =
        functionName === "TEXTJOIN" && coerceBooleanArg(retrievedArgs[1]),
      joinedValues = retrievedArgs
        .slice(functionName === "TEXTJOIN" ? 2 : 0)
        .flatMap(flattenNestedValues)
        .filter(
          (flatItem) =>
            !shouldIgnoreEmpty || (flatItem !== null && flatItem !== ""),
        );
    return (
      joinedValues.reduce(
        (totalLength, textItem) =>
          totalLength + coerceValueToText(textItem).length,
        0,
      ) +
        Math.max(0, joinedValues.length - 1) * delimiterText.length >
        32767 && raiseIssue("#LIMIT!"),
      joinedValues.map(coerceValueToText).join(delimiterText)
    );
  }
  if (functionName === "MAXIFS" || functionName === "MINIFS") {
    (retrievedArgs.length < 3 || retrievedArgs.length % 2 === 0) &&
      raiseIssue();
    const valueRange = flattenNestedValues(retrievedArgs[0]),
      criteriaPairs = [];
    for (
      let criteriaIndex = 1;
      criteriaIndex < retrievedArgs.length;
      criteriaIndex += 2
    ) {
      const criteriaRange = flattenNestedValues(retrievedArgs[criteriaIndex]);
      if (
        ((criteriaRange.length !== valueRange.length ||
          toMatrixArgument(retrievedArgs[criteriaIndex]).length !==
            toMatrixArgument(retrievedArgs[0]).length ||
          toMatrixArgument(retrievedArgs[criteriaIndex])[0].length !==
            toMatrixArgument(retrievedArgs[0])[0].length ||
          Array.isArray(retrievedArgs[criteriaIndex + 1])) &&
          raiseIssue(),
        evaluationContext.isIssue(retrievedArgs[criteriaIndex + 1]))
      )
        throw retrievedArgs[criteriaIndex + 1];
      criteriaPairs.push([criteriaRange, retrievedArgs[criteriaIndex + 1]]);
    }
    valueRange.length * criteriaPairs.length > 1e5 && raiseIssue("#LIMIT!");
    const matchingValues = valueRange.filter((g, valueIndex) =>
        criteriaPairs.every(
          ([rangeValues, criterion]) =>
            !evaluationContext.isIssue(rangeValues[valueIndex]) &&
            evaluationContext.matches(rangeValues[valueIndex], criterion),
        ),
      ),
      firstError = matchingValues.find(evaluationContext.isIssue);
    if (firstError) return firstError;
    const numericValues = matchingValues.filter(
      (matchedValue) => typeof matchedValue == "number",
    );
    return numericValues.length
      ? numericValues.reduce((extremum, candidateValue) =>
          functionName === "MAXIFS"
            ? Math.max(extremum, candidateValue)
            : Math.min(extremum, candidateValue),
        )
      : 0;
  }
  if (functionName === "XMATCH" || functionName === "XLOOKUP") {
    ((argRefs.length < (functionName === "XMATCH" ? 2 : 3) ||
      argRefs.length > (functionName === "XMATCH" ? 4 : 6)) &&
      raiseIssue(),
      (retrievedArgs[0] === undefined ||
        retrievedArgs[1] === undefined ||
        (functionName === "XLOOKUP" && retrievedArgs[2] === undefined)) &&
        raiseIssue());
    const lookupValue = retrievedArgs[0],
      matrixArgument = toMatrixArgument(retrievedArgs[1]);
    if (evaluationContext.isIssue(lookupValue)) return lookupValue;
    if (evaluationContext.isIssue(retrievedArgs[1])) return retrievedArgs[1];
    (Array.isArray(lookupValue) ||
      !matrixArgument.length ||
      !matrixArgument[0].length ||
      matrixArgument.some(
        (lookupRow) => lookupRow.length !== matrixArgument[0].length,
      ) ||
      (matrixArgument.length > 1 && matrixArgument[0].length > 1)) &&
      raiseIssue();
    const lookupValues = flattenNestedValues(matrixArgument),
      matchMode = coerceNumberArg(
        retrievedArgs[functionName === "XMATCH" ? 2 : 4],
        0,
      ),
      searchMode = coerceNumberArg(
        retrievedArgs[functionName === "XMATCH" ? 3 : 5],
        1,
      );
    (![0, -1, 1, 2].includes(matchMode) ||
      ![1, -1, 2, -2].includes(searchMode)) &&
      raiseIssue();
    const valueError = lookupValues.find(evaluationContext.isIssue);
    if (valueError) return valueError;
    let returnMatrix = [];
    if (functionName === "XLOOKUP") {
      if (evaluationContext.isIssue(retrievedArgs[2])) return retrievedArgs[2];
      ((returnMatrix = toMatrixArgument(retrievedArgs[2])),
        (!returnMatrix.length ||
          !returnMatrix[0].length ||
          returnMatrix.some(
            (returnRow) => returnRow.length !== returnMatrix[0].length,
          )) &&
          raiseIssue(),
        (matrixArgument.length > 1
          ? returnMatrix.length !== lookupValues.length
          : returnMatrix[0].length !== lookupValues.length) && raiseIssue());
    }
    (Math.abs(searchMode) === 2 || matchMode === -1 || matchMode === 1) &&
      (!["number", "string", "boolean"].includes(typeof lookupValue) ||
        lookupValues.some(
          (lookupItem) => typeof lookupItem != typeof lookupValue,
        )) &&
      raiseIssue();
    const searchOrder = lookupValues.map((A, sourceIndex) => sourceIndex);
    (searchMode < 0 && searchOrder.reverse(),
      Math.abs(searchMode) === 2 &&
        lookupValues.some(
          (sortedValue, sortedIndex) =>
            sortedIndex > 0 &&
            compareSortValues(sortedValue, lookupValues[sortedIndex - 1]) *
              Math.sign(searchMode) <
              0,
        ) &&
        raiseIssue());
    let comparisonCost = 0,
      matchIndex = searchOrder.find((candidateIndex) =>
        matchMode !== 2 || typeof lookupValue != "string"
          ? valuesEqualIgnoringCase(lookupValues[candidateIndex], lookupValue)
          : typeof lookupValues[candidateIndex] != "string"
            ? false
            : ((comparisonCost +=
                lookupValues[candidateIndex].length * lookupValue.length),
              comparisonCost > 1e5 && raiseIssue("#LIMIT!"),
              evaluationContext.matches(
                lookupValues[candidateIndex],
                `=${lookupValue}`,
              )),
      );
    if (matchIndex === undefined && (matchMode === -1 || matchMode === 1)) {
      const approximateMatches = searchOrder.filter(
        (approxIndex) =>
          Math.sign(
            compareSortValues(lookupValues[approxIndex], lookupValue),
          ) === matchMode,
      );
      (approximateMatches.sort(
        (leftIndex, rightIndex) =>
          compareSortValues(lookupValues[leftIndex], lookupValues[rightIndex]) *
          matchMode,
      ),
        (matchIndex = approximateMatches[0]));
    }
    return matchIndex === undefined
      ? functionName === "XLOOKUP" && retrievedArgs[3] !== undefined
        ? (retrievedArgs[3] ?? 0)
        : raiseIssue("#N/A")
      : functionName === "XMATCH"
        ? matchIndex + 1
        : matrixArgument.length > 1
          ? returnMatrix[matchIndex].length === 1
            ? (returnMatrix[matchIndex][0] ?? 0)
            : [returnMatrix[matchIndex]]
          : returnMatrix.length === 1
            ? (returnMatrix[0][matchIndex] ?? 0)
            : returnMatrix.map((returnMatrixRow) => [
                returnMatrixRow[matchIndex],
              ]);
  }
  if (functionName === "SEQUENCE") {
    retrievedArgs[0] === undefined &&
      retrievedArgs[1] === undefined &&
      raiseIssue();
    const sequenceRows = coerceIntegerArg(retrievedArgs[0], 1),
      sequenceColumns = coerceIntegerArg(retrievedArgs[1], 1),
      startValue = coerceNumberArg(retrievedArgs[2], 1),
      stepValue = coerceNumberArg(retrievedArgs[3], 1);
    return (
      validateArrayShape(sequenceRows, sequenceColumns),
      Array.from({ length: sequenceRows }, (v, sequenceRowIndex) =>
        Array.from(
          { length: sequenceColumns },
          (b, sequenceColumnIndex) =>
            startValue +
            (sequenceRowIndex * sequenceColumns + sequenceColumnIndex) *
              stepValue,
        ),
      )
    );
  }
  if (functionName === "FILTER") {
    const matrixArgument = toMatrixArgument(retrievedArgs[0]),
      includeMatrix = toMatrixArgument(retrievedArgs[1]);
    retrievedArgs[1] === undefined && raiseIssue();
    const maskError = flattenNestedValues(includeMatrix).find(
      evaluationContext.isIssue,
    );
    if (maskError) throw maskError;
    let filteredRows;
    if (
      includeMatrix.length === matrixArgument.length &&
      includeMatrix[0].length === 1
    )
      filteredRows = matrixArgument.filter((v, filterRowIndex) =>
        coerceBooleanArg(includeMatrix[filterRowIndex][0]),
      );
    else if (
      includeMatrix.length === 1 &&
      includeMatrix[0].length === matrixArgument[0].length
    )
      filteredRows = matrixArgument.map((filterSourceRow) =>
        filterSourceRow.filter((g, filterColumnIndex) =>
          coerceBooleanArg(includeMatrix[0][filterColumnIndex]),
        ),
      );
    else return raiseIssue();
    return !filteredRows.length || !filteredRows[0].length
      ? retrievedArgs[2] !== undefined
        ? toMatrixArgument(retrievedArgs[2])
        : raiseIssue("#CALC!")
      : filteredRows;
  }
  if (functionName === "TRANSPOSE")
    return toMatrixArgument(retrievedArgs[0])[0].map((h, transposeColumn) =>
      toMatrixArgument(retrievedArgs[0]).map(
        (transposeSourceRow) => transposeSourceRow[transposeColumn],
      ),
    );
  if (functionName === "SORT" || functionName === "SORTBY") {
    const sortableRows = toMatrixArgument(retrievedArgs[0]).map((rowCopy) => [
      ...rowCopy,
    ]);
    if (functionName === "SORTBY") {
      const matrixArgument = toMatrixArgument(retrievedArgs[1]);
      retrievedArgs[1] === undefined && raiseIssue();
      const isRowWiseSort =
          matrixArgument.length === 1 &&
          matrixArgument[0].length === sortableRows[0].length &&
          (sortableRows.length === 1 || matrixArgument[0].length > 1),
        sortItemCount = isRowWiseSort
          ? sortableRows[0].length
          : sortableRows.length,
        sortKeys = [];
      for (let keyIndex = 1; keyIndex < retrievedArgs.length; keyIndex += 2) {
        retrievedArgs[keyIndex] === undefined && raiseIssue();
        const matrixArgument = toMatrixArgument(retrievedArgs[keyIndex]),
          order = coerceNumberArg(retrievedArgs[keyIndex + 1], 1);
        ((isRowWiseSort
          ? matrixArgument.length !== 1 ||
            matrixArgument[0].length !== sortItemCount
          : matrixArgument.length !== sortItemCount ||
            matrixArgument.some((keyRow) => keyRow.length !== 1)) ||
          ![-1, 1].includes(order)) &&
          raiseIssue();
        const keyValues = flattenNestedValues(matrixArgument),
          keyError = keyValues.find(evaluationContext.isIssue);
        if (keyError) throw keyError;
        sortKeys.push({ data: keyValues, order: order });
      }
      const sortedIndexes = Array.from(
        { length: sortItemCount },
        (k, itemIndex) => itemIndex,
      ).sort((leftItemIndex, rightItemIndex) => {
        for (const sortKey of sortKeys) {
          const keyComparison = compareSortKeyWithDirection(
            sortKey.data[leftItemIndex],
            sortKey.data[rightItemIndex],
            sortKey.order,
          );
          if (keyComparison) return keyComparison;
        }
        return leftItemIndex - rightItemIndex;
      });
      return isRowWiseSort
        ? sortableRows.map((sortedRowValues) =>
            sortedIndexes.map(
              (sortedColumnIndex) => sortedRowValues[sortedColumnIndex],
            ),
          )
        : sortedIndexes.map((sortedRowIndex) => sortableRows[sortedRowIndex]);
    }
    const sortColumnIndex = coerceIntegerArg(retrievedArgs[1], 1) - 1,
      sortDirection = coerceNumberArg(retrievedArgs[2], 1),
      isSortByColumn =
        retrievedArgs[3] === undefined
          ? false
          : coerceBooleanArg(retrievedArgs[3]);
    [-1, 1].includes(sortDirection) || raiseIssue();
    const orientedRows = isSortByColumn
      ? sortableRows[0].map((b, transposeColumnIndex) =>
          sortableRows.map(
            (transposeRow) => transposeRow[transposeColumnIndex],
          ),
        )
      : sortableRows;
    (sortColumnIndex < 0 || sortColumnIndex >= orientedRows[0].length) &&
      raiseIssue();
    const keyColumnError = orientedRows
      .map((orientedRow) => orientedRow[sortColumnIndex])
      .find(evaluationContext.isIssue);
    if (keyColumnError) throw keyColumnError;
    return (
      orientedRows.sort((leftRow, rightRow) =>
        compareSortKeyWithDirection(
          leftRow[sortColumnIndex],
          rightRow[sortColumnIndex],
          sortDirection,
        ),
      ),
      isSortByColumn
        ? orientedRows[0].map((b, restoreColumnIndex) =>
            orientedRows.map((restoreRow) => restoreRow[restoreColumnIndex]),
          )
        : orientedRows
    );
  }
  if (functionName === "UNIQUE") {
    const matrixArgument = toMatrixArgument(retrievedArgs[0]),
      isByColumn =
        retrievedArgs[1] === undefined
          ? false
          : coerceBooleanArg(retrievedArgs[1]),
      isExactlyOnce =
        retrievedArgs[2] === undefined
          ? false
          : coerceBooleanArg(retrievedArgs[2]),
      rowsToCompare = isByColumn
        ? matrixArgument[0].map((w, uniqueColumnIndex) =>
            matrixArgument.map(
              (transposedRow) => transposedRow[uniqueColumnIndex],
            ),
          )
        : matrixArgument,
      rowKey = (keyInputRow) =>
        JSON.stringify(
          keyInputRow.map((keyCell) =>
            typeof keyCell == "string" ? keyCell.toLowerCase() : (keyCell ?? 0),
          ),
        ),
      occurrenceCounts = new Map();
    rowsToCompare.forEach((valueRow) =>
      occurrenceCounts.set(
        rowKey(valueRow),
        (occurrenceCounts.get(rowKey(valueRow)) ?? 0) + 1,
      ),
    );
    const seenKeys = new Set(),
      uniqueRows = rowsToCompare.filter((filterRow) => {
        const filterKey = rowKey(filterRow);
        return seenKeys.has(filterKey) ||
          (isExactlyOnce && occurrenceCounts.get(filterKey) !== 1)
          ? false
          : (seenKeys.add(filterKey), true);
      });
    return (
      uniqueRows.length || raiseIssue("#CALC!"),
      isByColumn
        ? uniqueRows[0].map((w, resultColumnIndex) =>
            uniqueRows.map(
              (uniqueResultRow) => uniqueResultRow[resultColumnIndex],
            ),
          )
        : uniqueRows
    );
  }
  if (functionName === "TAKE" || functionName === "DROP") {
    const matrixArgument = toMatrixArgument(retrievedArgs[0]),
      rowCount =
        retrievedArgs[1] === undefined
          ? undefined
          : coerceIntegerArg(retrievedArgs[1]),
      columnCount =
        retrievedArgs[2] === undefined
          ? undefined
          : coerceIntegerArg(retrievedArgs[2]),
      takeOrDropAxis = (axisItems, axisCount) =>
        axisCount === undefined
          ? axisItems
          : functionName === "TAKE"
            ? axisCount > 0
              ? axisItems.slice(0, axisCount)
              : axisItems.slice(axisCount)
            : axisCount >= 0
              ? axisItems.slice(axisCount)
              : axisItems.slice(0, axisCount);
    (rowCount === undefined && columnCount === undefined && raiseIssue(),
      (rowCount === 0 || columnCount === 0) && raiseIssue("#CALC!"));
    const result = takeOrDropAxis(matrixArgument, rowCount).map((resultRow) =>
      takeOrDropAxis(resultRow, columnCount),
    );
    return (
      (!result.length || !result[0].length) && raiseIssue("#CALC!"),
      result
    );
  }
  if (functionName === "CHOOSECOLS" || functionName === "CHOOSEROWS") {
    const matrixArgument = toMatrixArgument(retrievedArgs[0]),
      axisLength =
        functionName === "CHOOSECOLS"
          ? matrixArgument[0].length
          : matrixArgument.length,
      selectedIndexes = retrievedArgs
        .slice(1)
        .flatMap(flattenNestedValues)
        .map((indexArgument) => {
          const requestedIndex = coerceIntegerArg(indexArgument),
            zeroBasedIndex =
              requestedIndex > 0
                ? requestedIndex - 1
                : axisLength + requestedIndex;
          return (
            (!requestedIndex ||
              zeroBasedIndex < 0 ||
              zeroBasedIndex >= axisLength) &&
              raiseIssue(),
            zeroBasedIndex
          );
        });
    return (
      selectedIndexes.length || raiseIssue(),
      functionName === "CHOOSECOLS"
        ? matrixArgument.map((sourceMatrixRow) =>
            selectedIndexes.map(
              (selectedColumnIndex) => sourceMatrixRow[selectedColumnIndex],
            ),
          )
        : selectedIndexes.map(
            (selectedRowSource) => matrixArgument[selectedRowSource],
          )
    );
  }
  if (functionName === "HSTACK" || functionName === "VSTACK") {
    const stackedMatrices = retrievedArgs.map(toMatrixArgument),
      totalRows =
        functionName === "VSTACK"
          ? stackedMatrices.reduce(
              (rowTotal, matrixForRows) => rowTotal + matrixForRows.length,
              0,
            )
          : Math.max(
              ...stackedMatrices.map(
                (matrixForRowCount) => matrixForRowCount.length,
              ),
            ),
      totalColumns =
        functionName === "HSTACK"
          ? stackedMatrices.reduce(
              (columnTotal, matrixForColumns) =>
                columnTotal + matrixForColumns[0].length,
              0,
            )
          : Math.max(
              ...stackedMatrices.map(
                (matrixForColumnCount) => matrixForColumnCount[0].length,
              ),
            );
    return (
      validateArrayShape(totalRows, totalColumns),
      functionName === "VSTACK"
        ? stackedMatrices.flatMap((sourceMatrix) =>
            sourceMatrix.map((stackedSourceRow) =>
              Array.from({ length: totalColumns }, (g, columnIndex) =>
                columnIndex < stackedSourceRow.length
                  ? stackedSourceRow[columnIndex]
                  : evaluationContext.issue("#N/A"),
              ),
            ),
          )
        : Array.from({ length: totalRows }, (m, rowIndex) =>
            stackedMatrices.flatMap((stackedMatrix) =>
              Array.from(
                { length: stackedMatrix[0].length },
                (b, stackedColumnIndex) =>
                  rowIndex < stackedMatrix.length
                    ? stackedMatrix[rowIndex][stackedColumnIndex]
                    : evaluationContext.issue("#N/A"),
              ),
            ),
          )
    );
  }
  if (functionName === "TOCOL" || functionName === "TOROW") {
    const matrixArgument = toMatrixArgument(retrievedArgs[0]),
      ignoreMode = coerceIntegerArg(retrievedArgs[1], 0),
      scanByRow =
        retrievedArgs[2] === undefined
          ? false
          : coerceBooleanArg(retrievedArgs[2]);
    [0, 1, 2, 3].includes(ignoreMode) || raiseIssue();
    const flatValues = (
      scanByRow
        ? matrixArgument[0].flatMap((v, sourceColumnIndex) =>
            matrixArgument.map(
              (sourceRowCells) => sourceRowCells[sourceColumnIndex],
            ),
          )
        : matrixArgument.flat()
    ).filter(
      (flatValue) =>
        !((ignoreMode === 1 || ignoreMode === 3) && flatValue === null) &&
        !(
          (ignoreMode === 2 || ignoreMode === 3) &&
          evaluationContext.isIssue(flatValue)
        ),
    );
    return (
      flatValues.length || raiseIssue("#CALC!"),
      functionName === "TOCOL"
        ? flatValues.map((columnOutputValue) => [columnOutputValue ?? 0])
        : [flatValues.map((rowOutputValue) => rowOutputValue ?? 0)]
    );
  }
  if (functionName === "EXPAND") {
    const matrixArgument = toMatrixArgument(retrievedArgs[0]),
      expandedRows = coerceIntegerArg(retrievedArgs[1], matrixArgument.length),
      expandedColumns = coerceIntegerArg(
        retrievedArgs[2],
        matrixArgument[0].length,
      );
    return (
      validateArrayShape(expandedRows, expandedColumns),
      (expandedRows < matrixArgument.length ||
        expandedColumns < matrixArgument[0].length) &&
        raiseIssue(),
      Array.from({ length: expandedRows }, (m, expandRowIndex) =>
        Array.from({ length: expandedColumns }, (g, expandColumnIndex) =>
          expandRowIndex < matrixArgument.length &&
          expandColumnIndex < matrixArgument[expandRowIndex].length
            ? matrixArgument[expandRowIndex][expandColumnIndex]
            : (retrievedArgs[3] ?? evaluationContext.issue("#N/A")),
        ),
      )
    );
  }
  if (functionName === "TEXTBEFORE" || functionName === "TEXTAFTER") {
    const textValue = coerceValueToText(retrievedArgs[0]),
      delimiterValue = coerceValueToText(retrievedArgs[1]),
      instanceNumber = coerceIntegerArg(retrievedArgs[2], 1),
      isCaseInsensitive = coerceIntegerArg(retrievedArgs[3], 0) === 1,
      shouldMatchEnd = coerceIntegerArg(retrievedArgs[4], 0) === 1;
    (!instanceNumber || !delimiterValue) && raiseIssue();
    const searchText = isCaseInsensitive ? textValue.toLowerCase() : textValue,
      searchDelimiter = isCaseInsensitive
        ? delimiterValue.toLowerCase()
        : delimiterValue,
      delimiterPositions = [];
    for (
      let searchStart = 0;
      searchStart <= searchText.length - searchDelimiter.length;
    ) {
      const foundIndex = searchText.indexOf(searchDelimiter, searchStart);
      if (foundIndex < 0) break;
      (delimiterPositions.push(foundIndex),
        (searchStart = foundIndex + searchDelimiter.length));
    }
    const matchPosition =
      delimiterPositions[
        instanceNumber > 0
          ? instanceNumber - 1
          : delimiterPositions.length + instanceNumber
      ];
    return matchPosition === undefined
      ? shouldMatchEnd &&
        Math.abs(instanceNumber) === delimiterPositions.length + 1
        ? functionName === "TEXTBEFORE"
          ? instanceNumber > 0
            ? textValue
            : ""
          : instanceNumber > 0
            ? ""
            : textValue
        : retrievedArgs[5] !== undefined
          ? retrievedArgs[5]
          : raiseIssue("#N/A")
      : functionName === "TEXTBEFORE"
        ? textValue.slice(0, matchPosition)
        : textValue.slice(matchPosition + delimiterValue.length);
  }
  if (functionName === "TEXTSPLIT") {
    (retrievedArgs.length < 2 || retrievedArgs.length > 6) && raiseIssue();
    const argumentError = retrievedArgs
      .slice(0, 5)
      .flatMap(flattenNestedValues)
      .find(evaluationContext.isIssue);
    if (argumentError) return argumentError;
    const textToSplit = coerceValueToText(retrievedArgs[0]),
      shouldIgnoreEmptyValues =
        retrievedArgs[3] === undefined
          ? false
          : coerceBooleanArg(retrievedArgs[3]),
      textMatchMode = coerceNumberArg(retrievedArgs[4], 0);
    (Array.isArray(retrievedArgs[0]) || ![0, 1].includes(textMatchMode)) &&
      raiseIssue();
    const toTextValues = (valueArgument) =>
        flattenNestedValues(valueArgument)
          .filter((nonNullItem) => nonNullItem != null)
          .map((splitItem) => {
            if (evaluationContext.isIssue(splitItem)) throw splitItem;
            return coerceValueToText(splitItem);
          })
          .filter(Boolean),
      columnDelimiters = toTextValues(retrievedArgs[1]),
      rowDelimiters = toTextValues(retrievedArgs[2]);
    (!columnDelimiters.length && !rowDelimiters.length && raiseIssue(),
      (columnDelimiters.length + rowDelimiters.length > 100 ||
        [...columnDelimiters, ...rowDelimiters].join("").length > 1e4) &&
        raiseIssue("#LIMIT!"));
    const splitByDelimiters = (splitText, delimiterList) =>
        delimiterList.length
          ? splitText
              .split(
                new RegExp(
                  delimiterList
                    .map((delimiterItem) =>
                      delimiterItem.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"),
                    )
                    .join("|"),
                  textMatchMode ? "i" : "",
                ),
              )
              .filter((segment) => !shouldIgnoreEmptyValues || segment !== "")
          : [splitText],
      splitRows = splitByDelimiters(textToSplit, rowDelimiters).map((rowText) =>
        splitByDelimiters(rowText, columnDelimiters),
      ),
      maxColumnCount = Math.max(
        ...splitRows.map((splitRowForWidth) => splitRowForWidth.length),
      );
    return (
      validateArrayShape(splitRows.length, maxColumnCount),
      splitRows.map((splitRowSegments) =>
        Array.from(
          { length: maxColumnCount },
          (k, segmentIndex) =>
            splitRowSegments[segmentIndex] ??
            retrievedArgs[5] ??
            evaluationContext.issue("#N/A"),
        ),
      )
    );
  }
  return notModernFunctionSentinel;
}
function validateDefinedName(definedNames, workbook) {
  if (definedNames === undefined) return [];
  if (!Array.isArray(definedNames) || definedNames.length > 1e3)
    throw new Error("A workbook may contain at most 1,000 defined names.");
  const seenScopeKeys = new Set();
  return definedNames.map((definedName) => {
    if (
      !definedName ||
      typeof definedName != "object" ||
      Array.isArray(definedName)
    )
      throw new Error("Invalid defined name.");
    const entry = definedName;
    if (
      Object.keys(entry).some(
        (key) => !["name", "refersTo", "scopeSheetId", "comment"].includes(key),
      )
    )
      throw new Error("Unknown defined-name field.");
    if (
      typeof entry.name != "string" ||
      !/^[A-Za-z_\\][A-Za-z0-9_.\\]{0,254}$/.test(entry.name) ||
      /^(?:[A-Za-z]{1,3}[1-9]\d*|R\d+C\d+|R|C)$/i.test(entry.name)
    )
      throw new Error(
        "Names must begin with a letter or underscore and cannot be cell addresses.",
      );
    if (
      typeof entry.refersTo != "string" ||
      !entry.refersTo.startsWith("=") ||
      entry.refersTo.length > 4096 ||
      entry.refersTo.length < 2
    )
      throw new Error(
        "A name must refer to a formula beginning with = (up to 4,096 characters).",
      );
    if (
      entry.scopeSheetId !== undefined &&
      (typeof entry.scopeSheetId != "string" ||
        !workbook.sheets.some((sheet) => sheet.id === entry.scopeSheetId))
    )
      throw new Error("Defined-name worksheet scope was not found.");
    if (
      entry.comment !== undefined &&
      (typeof entry.comment != "string" || entry.comment.length > 1e3)
    )
      throw new Error(
        "Defined-name comments may contain at most 1,000 characters.",
      );
    const scopeKey = `${entry.scopeSheetId ?? ""}:${entry.name.toLowerCase()}`;
    if (seenScopeKeys.has(scopeKey))
      throw new Error("A name already exists in this scope.");
    return (
      seenScopeKeys.add(scopeKey),
      {
        name: entry.name,
        refersTo: entry.refersTo,
        ...(entry.scopeSheetId ? { scopeSheetId: entry.scopeSheetId } : {}),
        ...(entry.comment ? { comment: entry.comment } : {}),
      }
    );
  });
}
const definedNameScopeKey = (definedName) =>
    `${definedName.scopeSheetId ?? ""}:${definedName.name.toLowerCase()}`,
  formatQuotedSheetPrefix = (sheetName) =>
    `'${sheetName.replace(/'/g, "''")}'!`;
function inspectFormulaForMigration(formulaText, throwOnError = false) {
  const references = [];
  let end = 0;
  const reportIssue = (message) => {
      if (throwOnError) throw Error(message);
    },
    skipTableReference = () => {
      let depth = 0;
      do
        (formulaText[end] === "[" && depth++,
          formulaText[end] === "]" && depth--,
          end++);
      while (end < formulaText.length && depth);
      depth &&
        reportIssue("Unbalanced table references cannot be safely migrated.");
    };
  for (; end < formulaText.length;) {
    const start = end,
      currentChar = formulaText[end];
    if (currentChar === '"') {
      end++;
      let isTerminated = false;
      for (; end < formulaText.length;)
        if (formulaText[end++] === '"')
          if (formulaText[end] === '"') end++;
          else {
            isTerminated = true;
            break;
          }
      isTerminated ||
        reportIssue(
          "An unterminated formula string cannot be safely migrated.",
        );
      continue;
    }
    if (currentChar === "#") {
      const errorMatch =
        /^#(?:REF!|NAME\?|VALUE!|DIV\/0!|N\/A|NUM!|NULL!|SPILL!|CALC!|CYCLE!|LIMIT!|BLOCKED!|ERROR!)/i.exec(
          formulaText.slice(end),
        );
      if (errorMatch) {
        end += errorMatch[0].length;
        continue;
      }
      end++;
      continue;
    }
    const threeDimMatch =
      /^(?:'(?:[^']|'')+'|[A-Za-z0-9_.]+):(?:'(?:[^']|'')+'|[A-Za-z0-9_.]+)!/.exec(
        formulaText.slice(end),
      );
    if (threeDimMatch) {
      (reportIssue("Three-dimensional references cannot be safely migrated."),
        (end += threeDimMatch[0].length));
      continue;
    }
    let qualifier;
    if (currentChar === "'") {
      const quotedMatch = /^'((?:[^']|'')+)'!/.exec(formulaText.slice(end));
      if (!quotedMatch) {
        for (
          reportIssue(
            "An unsupported quoted reference cannot be safely migrated.",
          ),
            end++;
          end < formulaText.length && formulaText[end++] !== "'";
        );
        continue;
      }
      ((qualifier = quotedMatch[1].replace(/''/g, "'")),
        (end += quotedMatch[0].length),
        /[\[\]:|]/.test(qualifier) &&
          reportIssue(
            "External or three-dimensional references cannot be safely migrated.",
          ));
    } else {
      const unquotedMatch = /^([A-Za-z0-9_\\][A-Za-z0-9_.\\]*)!/.exec(
        formulaText.slice(end),
      );
      unquotedMatch &&
        ((qualifier = unquotedMatch[1]), (end += unquotedMatch[0].length));
    }
    if (qualifier === undefined && currentChar === "[") {
      (formulaText.startsWith("[@", end) ||
        reportIssue(
          "External or unqualified bracket references cannot be safely migrated.",
        ),
        skipTableReference());
      continue;
    }
    const cellRangeMatch =
      /^(?:\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}|\$?[1-9]\d*:\$?[1-9]\d*)(?![A-Za-z0-9_.\\])/.exec(
        formulaText.slice(end),
      );
    if (cellRangeMatch) {
      ((end += cellRangeMatch[0].length),
        references.push({
          start: start,
          end: end,
          text: formulaText.slice(start, end),
          name: cellRangeMatch[0],
          ...(qualifier ? { qualifier: qualifier } : {}),
          cell: true,
        }));
      continue;
    }
    if (qualifier === undefined && /[0-9.]/.test(currentChar)) {
      const numberMatch = /^(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?/.exec(
        formulaText.slice(end),
      );
      if (numberMatch) {
        end += numberMatch[0].length;
        continue;
      }
    }
    const nameMatch =
      /^(?:\$?[A-Za-z]{1,3}\$?[1-9]\d*(?![A-Za-z0-9_.\\])|[A-Za-z_\\][A-Za-z0-9_.\\]*)/.exec(
        formulaText.slice(end),
      );
    if (!nameMatch) {
      (qualifier &&
        reportIssue("The qualified reference cannot be safely migrated."),
        currentChar === "!" &&
          reportIssue("The reference qualifier cannot be safely migrated."),
        currentChar === "@" && formulaText[end + 1] === "["
          ? (end++, skipTableReference())
          : end++);
      continue;
    }
    const nameText = nameMatch[0];
    if (((end += nameText.length), formulaText[end] === "[")) {
      skipTableReference();
      continue;
    }
    const isFunctionCall = /^\s*\(/.test(formulaText.slice(end));
    (isFunctionCall &&
      /^(?:_xlfn\.)?(?:INDIRECT|EVALUATE|GET\.CELL)$/i.test(nameText) &&
      reportIssue(
        `${nameText} uses dynamic or locally bound names and cannot be safely migrated.`,
      ),
      !(start > 0 && /[A-Za-z0-9_.$\\]/.test(formulaText[start - 1])) &&
        references.push({
          start: start,
          end: end,
          text: formulaText.slice(start, end),
          name: nameText,
          ...(isFunctionCall ? { call: true } : {}),
          ...(qualifier ? { qualifier: qualifier } : {}),
          ...(/^\$?[A-Za-z]{1,3}\$?[1-9]\d*$/.test(nameText)
            ? { cell: true }
            : {}),
        }));
  }
  return references;
}
function m5(formulaText, throwOnError = false) {
  const references = inspectFormulaForMigration(formulaText, throwOnError);
  if (!/\b(?:_xlfn\.)?(?:LET|LAMBDA)\s*\(/i.test(formulaText))
    return references;
  const boundByStart = new Map();
  try {
    walkFormulaNodes(
      tokenizeFormulaText(formulaText),
      (token, boundNames, s, isExcluded) => {
        boundByStart.set(token.start, {
          bound: new Set(boundNames),
          excluded:
            isExcluded ||
            (!token.qualifier &&
              boundNames.has(normalizeBindingName(token.name))),
        });
      },
    );
  } catch (error) {
    if (throwOnError) throw error;
    return references;
  }
  return references
    .filter((candidate) => !boundByStart.get(candidate.start)?.excluded)
    .map((reference) => ({
      ...reference,
      bound: boundByStart.get(reference.start)?.bound,
    }));
}
function replaceTextSpans(text, spans, getReplacement) {
  let result = "",
    lastEnd = 0;
  for (const span of spans)
    ((result += text.slice(lastEnd, span.start) + getReplacement(span)),
      (lastEnd = span.end));
  return result + text.slice(lastEnd);
}
const pO = 500;
function stripFormulaStringLiterals(formulaText) {
  let output = "",
    index = 0;
  for (; index < formulaText.length;) {
    const currentChar = formulaText[index];
    if (currentChar === '"' || currentChar === "'") {
      const quoteStart = index++,
        quoteChar = currentChar;
      for (; index < formulaText.length;)
        if (formulaText[index++] === quoteChar)
          if (formulaText[index] === quoteChar) index++;
          else break;
      output += formulaText.slice(quoteStart, index);
      continue;
    }
    const bracketMatch = /^(?:[A-Za-z_][A-Za-z0-9_.]*\[|\[@)/.exec(
      formulaText.slice(index),
    );
    if (
      bracketMatch &&
      (index === 0 || !/[A-Za-z0-9_.$]/.test(formulaText[index - 1]))
    ) {
      const bracketStart = index;
      index += bracketMatch[0].length;
      let bracketDepth = 1;
      for (; index < formulaText.length && bracketDepth;)
        (formulaText[index] === "[" && bracketDepth++,
          formulaText[index] === "]" && bracketDepth--,
          index++);
      const bracketText = formulaText.slice(bracketStart, index);
      !bracketDepth &&
      !/[!|]/.test(bracketText) &&
      !/^[A-Za-z0-9_!]/.test(formulaText.slice(index))
        ? (output += "0")
        : (output += bracketText);
      continue;
    }
    output += formulaText[index++];
  }
  return output;
}
function $v(
  formulaText,
  oldTableName,
  newTableName,
  names = {},
  canMatchThisRow = false,
) {
  const columnRenames = new Map(
    Object.entries(names).map(([originalName, newName]) => [
      originalName.toLowerCase(),
      newName,
    ]),
  );
  let result = "",
    index = 0;
  for (; index < formulaText.length;) {
    const currentChar = formulaText[index];
    if (currentChar === '"' || currentChar === "'") {
      const quoteStart = index++,
        quoteChar = currentChar;
      for (; index < formulaText.length;)
        if (formulaText[index++] === quoteChar)
          if (formulaText[index] === quoteChar) index++;
          else break;
      result += formulaText.slice(quoteStart, index);
      continue;
    }
    const tableRefMatch = /^([A-Za-z_][A-Za-z0-9_.]*)?\[/.exec(
      formulaText.slice(index),
    );
    if (
      tableRefMatch &&
      (index === 0 || !/[A-Za-z0-9_.$]/.test(formulaText[index - 1]))
    ) {
      const refStart = index;
      index += tableRefMatch[0].length;
      let bracketDepth = 1;
      for (; index < formulaText.length && bracketDepth;)
        (formulaText[index] === "[" && bracketDepth++,
          formulaText[index] === "]" && bracketDepth--,
          index++);
      const refText = formulaText.slice(refStart, index),
        qualifier = tableRefMatch[1] ?? "",
        isSameTable = qualifier
          ? qualifier.toLowerCase() === oldTableName.toLowerCase()
          : canMatchThisRow && refText.startsWith("[@");
      if (
        !bracketDepth &&
        isSameTable &&
        !/[!|]/.test(refText) &&
        !/^[A-Za-z0-9_!]/.test(formulaText.slice(index))
      ) {
        const rewrittenRef = refText
          .slice(qualifier.length)
          .replace(
            /\[(@?)([^\[\]]+)\]/g,
            (wholeMatch, atMarker, columnName) => {
              const newColumnName = columnRenames.get(columnName.toLowerCase());
              return newColumnName
                ? `[${atMarker}${newColumnName}]`
                : wholeMatch;
            },
          );
        result += `${qualifier ? newTableName : ""}${rewrittenRef}`;
      } else result += refText;
      continue;
    }
    result += formulaText[index++];
  }
  return result;
}
const reservedDefinedNames = new Set([
    "_xlnm.print_area",
    "_xlnm.print_titles",
    "_xlnm._filterdatabase",
  ]),
  spreadsheetMlNamespace =
    "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
  throwXlsxDefinedNameError = (reason) => {
    throw Error(
      `XLSX defined names: ${reason} Import/export stopped to prevent changed formula bindings.`,
    );
  },
  invalidXmlCharPattern = new RegExp(
    "[\\u0000-\\u0008\\u000B\\u000C\\u000E-\\u001F\\uFFFE\\uFFFF]|[\\uD800-\\uDBFF](?![\\uDC00-\\uDFFF])|(?<![\\uD800-\\uDBFF])[\\uDC00-\\uDFFF]",
  );
function nP(xmlText) {
  return xmlText.includes("<") ||
    xmlText.includes("]]>") ||
    invalidXmlCharPattern.test(xmlText)
    ? throwXlsxDefinedNameError("Malformed XML text.")
    : xmlText.replace(/&[^;\s<]*;|&/g, (entityText) => {
        const namedEntities = {
          "&amp;": "&",
          "&lt;": "<",
          "&gt;": ">",
          "&quot;": '"',
          "&apos;": "'",
        };
        if (namedEntities[entityText]) return namedEntities[entityText];
        if (!/^&#(?:[0-9]+|x[0-9a-fA-F]+);$/.test(entityText))
          return throwXlsxDefinedNameError("Unknown or malformed XML entity.");
        const isHexEntity = entityText.startsWith("&#x"),
          codePoint = Number.parseInt(
            entityText.slice(isHexEntity ? 3 : 2, -1),
            isHexEntity ? 16 : 10,
          );
        if (
          !Number.isSafeInteger(codePoint) ||
          codePoint > 1114111 ||
          codePoint < 0 ||
          (codePoint >= 55296 && codePoint <= 57343)
        )
          return throwXlsxDefinedNameError("Invalid XML character entity.");
        const character = String.fromCodePoint(codePoint);
        return invalidXmlCharPattern.test(character)
          ? throwXlsxDefinedNameError("Invalid XML character entity.")
          : character;
      });
}
function validateXmlSafeName(nameValue, shouldEscapeAsNumeric = false) {
  return invalidXmlCharPattern.test(nameValue)
    ? throwXlsxDefinedNameError("A name contains an invalid XML character.")
    : nameValue.replace(
        /[&<>"'\r\n\t]/g,
        (matchedChar) =>
          ({
            "&": "&amp;",
            "<": "&lt;",
            ">": "&gt;",
            '"': "&quot;",
            "'": "&apos;",
          })[matchedChar] ??
          (shouldEscapeAsNumeric
            ? `&#${matchedChar.charCodeAt(0)};`
            : matchedChar === "\r"
              ? "&#13;"
              : matchedChar),
      );
}
function parseDefinedNameXmlAttributes(attributeText) {
  const attributes = Object.create(null);
  let remaining = attributeText;
  for (; remaining.trim();) {
    const attributeMatch =
      /^\s+([A-Za-z_][A-Za-z0-9_.:-]*)\s*=\s*(["'])([\s\S]*?)\2/.exec(
        remaining,
      );
    if (!attributeMatch || Object.hasOwn(attributes, attributeMatch[1]))
      return throwXlsxDefinedNameError(
        "Malformed or duplicate defined-name XML attribute.",
      );
    ((attributes[attributeMatch[1]] = nP(
      attributeMatch[3].replace(/\r\n?|\n|\t/g, " "),
    )),
      (remaining = remaining.slice(attributeMatch[0].length)));
  }
  return attributes;
}
function parseDefinedNamesXml(workbookXml) {
  if (/<!DOCTYPE|<!ENTITY/i.test(workbookXml))
    return throwXlsxDefinedNameError(
      "DTD and entity declarations are not supported.",
    );
  const definedNamesMatches = [
    ...workbookXml.matchAll(
      /<((?:[A-Za-z_][\w.-]*:)?definedNames)\b((?:[^"'<>]|"[^"]*"|'[^']*')*?)(?:\/>|>([\s\S]*?)<\/\1\s*>)/g,
    ),
  ];
  if (definedNamesMatches.length > 1)
    return throwXlsxDefinedNameError("Duplicate definedNames sections.");
  const definedNamesMatch = definedNamesMatches[0],
    xmlWithoutDefinedNames = definedNamesMatch
      ? workbookXml.slice(0, definedNamesMatch.index) +
        workbookXml.slice(definedNamesMatch.index + definedNamesMatch[0].length)
      : workbookXml;
  if (/<\/?(?:[A-Za-z_][\w.-]*:)?definedNames?\b/.test(xmlWithoutDefinedNames))
    return throwXlsxDefinedNameError(
      "Malformed or misplaced defined-name elements.",
    );
  if (!definedNamesMatch) return;
  const definedNameXmlAttributes = parseDefinedNameXmlAttributes(
      definedNamesMatch[2],
    ),
    workbookTagMatch =
      /<(?:[A-Za-z_][\w.-]*:)?workbook\b((?:[^"'<>]|"[^"]*"|'[^']*')*?)>/.exec(
        workbookXml,
      ),
    namespaces = {
      ...(workbookTagMatch
        ? parseDefinedNameXmlAttributes(workbookTagMatch[1])
        : {}),
      ...definedNameXmlAttributes,
    },
    prefix = definedNamesMatch[1].slice(0, -12);
  if (
    namespaces[prefix ? `xmlns:${prefix.slice(0, -1)}` : "xmlns"] !==
    spreadsheetMlNamespace
  )
    return throwXlsxDefinedNameError(
      "Unsupported or undeclared definedNames namespace.",
    );
  for (const definedNameXmlAttribute of Object.keys(definedNameXmlAttributes))
    if (
      !definedNameXmlAttribute.startsWith("xmlns") ||
      definedNameXmlAttributes[definedNameXmlAttribute] !==
        spreadsheetMlNamespace
    )
      return throwXlsxDefinedNameError(
        "Unsupported definedNames namespace or attribute.",
      );
  return {
    start: definedNamesMatch.index,
    end: definedNamesMatch.index + definedNamesMatch[0].length,
    prefix: definedNamesMatch[1].slice(0, -12),
    body: definedNamesMatch[3] ?? "",
    namespaces: namespaces,
  };
}
function validateNameFormulaText(refersTo) {
  if (
    !refersTo.startsWith("=") ||
    refersTo.length < 2 ||
    refersTo.length > 4096
  )
    return throwXlsxDefinedNameError(
      "A definition must contain a formula of 1–4,095 characters after =.",
    );
  let parenDepth = 0,
    bracketDepth = 0,
    braceDepth = 0,
    position = 1,
    formulaBody = "";
  for (; position < refersTo.length;) {
    const currentChar = refersTo[position];
    if (!bracketDepth && (currentChar === '"' || currentChar === "'")) {
      const quoteChar = currentChar,
        quoteStart = position++;
      let isClosed = false;
      for (; position < refersTo.length;)
        if (refersTo[position++] === quoteChar)
          if (refersTo[position] === quoteChar) position++;
          else {
            isClosed = true;
            break;
          }
      if (!isClosed)
        return throwXlsxDefinedNameError(
          "Unterminated formula string or worksheet qualifier.",
        );
      if (quoteChar === "'") {
        if (refersTo[position] !== "!")
          return throwXlsxDefinedNameError(
            "Malformed quoted worksheet qualifier.",
          );
        formulaBody += refersTo.slice(quoteStart, position);
      } else formulaBody += '""';
      continue;
    }
    if (
      (currentChar === "[" && bracketDepth++,
      currentChar === "]" && --bracketDepth < 0)
    )
      return throwXlsxDefinedNameError("Unbalanced bracket reference.");
    if (!bracketDepth) {
      if (
        (currentChar === "{" && braceDepth++,
        currentChar === "}" && --braceDepth < 0)
      )
        return throwXlsxDefinedNameError("Unbalanced array braces.");
      if (
        (currentChar === "(" && parenDepth++,
        currentChar === ")" && --parenDepth < 0)
      )
        return throwXlsxDefinedNameError("Unbalanced formula parentheses.");
    }
    ((formulaBody += currentChar), position++);
  }
  if (parenDepth || bracketDepth || braceDepth)
    return throwXlsxDefinedNameError(
      "Unbalanced formula parentheses, array braces or bracket reference.",
    );
  const strippedText = stripFormulaStringLiterals(formulaBody);
  if (
    /[\[\]|]/.test(strippedText) ||
    /(?:https?:|file:|javascript:|data:|\\\\)/i.test(strippedText) ||
    /\b(?:_xlfn\.)?(?:WEBSERVICE|HYPERLINK|IMAGE|RTD|DDE|CALL|EXEC|EXECUTE|REGISTER(?:\.ID)?|RUN|SQL\.REQUEST|STOCKHISTORY|EVALUATE|GET\.CELL)\s*\(/i.test(
      strippedText,
    )
  )
    return throwXlsxDefinedNameError(
      "External or executable name formulas are not supported.",
    );
  if (
    /(?:'(?:[^']|'')+'|[A-Za-z0-9_.]+):(?:'(?:[^']|'')+'|[A-Za-z0-9_.]+)!/.test(
      strippedText,
    )
  )
    return throwXlsxDefinedNameError(
      "Three-dimensional name references are not supported.",
    );
  if (!formulaBody.trim())
    return throwXlsxDefinedNameError("An empty formula is not supported.");
}
function validateNameSheetReferences(refersTo, workbook) {
  const strippedRefersTo = stripFormulaStringLiterals(
    refersTo.replace(/"(?:[^"]|"")*"/g, '""'),
  );
  for (const sheetRefMatch of strippedRefersTo.matchAll(
    /(?:'((?:[^']|'')+)'|([A-Za-z_\\][A-Za-z0-9_.\\]*))!/g,
  )) {
    const sheetName = sheetRefMatch[1]?.replace(/''/g, "'") ?? sheetRefMatch[2];
    if (
      !workbook.sheets.some(
        (sheet) => sheet.name.toLowerCase() === sheetName.toLowerCase(),
      )
    )
      return throwXlsxDefinedNameError(
        `A name refers to an unknown worksheet "${sheetName}".`,
      );
  }
}
function insertDefinedNamesIntoWorkbookXml(workbookXml, builtinNameXml) {
  const definedNamesXml = parseDefinedNamesXml(workbookXml),
    definedNamesBlockXml = builtinNameXml
      ? `<definedNames xmlns="${spreadsheetMlNamespace}">${builtinNameXml}</definedNames>`
      : "";
  if (definedNamesXml)
    return (
      workbookXml.slice(0, definedNamesXml.start) +
      definedNamesBlockXml +
      workbookXml.slice(definedNamesXml.end)
    );
  if (!builtinNameXml) return workbookXml;
  const insertionMatch =
    /<(?:[A-Za-z_][\w.-]*:)?(?:calcPr|oleSize|customWorkbookViews|pivotCaches|smartTagPr|smartTagTypes|webPublishing|fileRecoveryPr|webPublishObjects|extLst)\b|<\/(?:[A-Za-z_][\w.-]*:)?workbook\s*>/.exec(
      workbookXml,
    );
  return insertionMatch
    ? workbookXml.slice(0, insertionMatch.index) +
        definedNamesBlockXml +
        workbookXml.slice(insertionMatch.index)
    : throwXlsxDefinedNameError(
        "The workbook XML has no supported insertion point.",
      );
}
const localBindingFunctionNames = ["LET", "LAMBDA"],
  reusableFormulaLimits = {
    depth: 64,
    calls: 1024,
    work: 1e6,
    retainedValues: 1e5,
  },
  isFormulaClosure = (value) =>
    !!value &&
    typeof value == "object" &&
    "kind" in value &&
    value.kind === "formula-closure",
  y5 = "__LGVALUE";
function evaluateReusableFormula(
  formulaText,
  workbook,
  sheetId,
  evaluationContext,
) {
  let workUnits = 0,
    callCount = 0,
    retainedValueCount = 0;
  const tokenCache = new Map(),
    activeNameKeys = new Set(),
    raiseCalcError = (errorCode, errorMessage) => {
      throw (
        evaluationContext.diagnostic(errorCode, errorMessage),
        evaluationContext.issue(errorCode)
      );
    },
    chargeWork = (workUnitCount, nestingDepth) => {
      ((workUnits += workUnitCount),
        (nestingDepth > reusableFormulaLimits.depth ||
          workUnits > reusableFormulaLimits.work) &&
          raiseCalcError(
            "#LIMIT!",
            `Reusable formula evaluation exceeds ${reusableFormulaLimits.depth} levels or ${reusableFormulaLimits.work.toLocaleString()} work units.`,
          ));
    },
    retainValues = (retainedValue) => {
      const retainedCellCount = Array.isArray(retainedValue)
        ? retainedValue.reduce(
            (total, row) => total + (Array.isArray(row) ? row.length : 1),
            0,
          )
        : 1;
      return (
        (retainedValueCount += retainedCellCount),
        retainedValueCount > reusableFormulaLimits.retainedValues &&
          raiseCalcError(
            "#LIMIT!",
            "Reusable formula intermediate values exceed 100,000 cells.",
          ),
        retainedValue
      );
    },
    cachedTokenize = (sourceText) => {
      let cachedTokens = tokenCache.get(sourceText);
      if (!cachedTokens) {
        chargeWork(sourceText.length, 0);
        try {
          try {
            validateNameFormulaText(
              sourceText.startsWith("=") ? sourceText : `=${sourceText}`,
            );
          } catch (error) {
            const validationMessage =
              error instanceof Error ? error.message : "Unsafe formula.";
            raiseCalcError(
              /External|executable|Three-dimensional/i.test(validationMessage)
                ? "#BLOCKED!"
                : "#VALUE!",
              validationMessage,
            );
          }
          ((cachedTokens = tokenizeFormulaText(sourceText)),
            walkFormulaNodes(cachedTokens, (formulaNode, F, I, O) => {
              (!O &&
                /^__LG/i.test(formulaNode.name) &&
                raiseCalcError(
                  "#NAME?",
                  "Internal calculation names cannot be called from a workbook.",
                ),
                /^(?:_xlfn\.)?(?:WEBSERVICE|HYPERLINK|IMAGE|RTD|DDE|CALL|EXEC|EXECUTE|REGISTER(?:\.ID)?|RUN|SQL\.REQUEST|STOCKHISTORY|EVALUATE|GET\.CELL)$/i.test(
                  formulaNode.name,
                ) &&
                  raiseCalcError(
                    "#BLOCKED!",
                    "External or executable custom functions are not supported.",
                  ));
            }));
        } catch (error) {
          if (evaluationContext.isIssue(error)) throw error;
          raiseCalcError(
            error instanceof Error && error.message === "#LIMIT!"
              ? "#LIMIT!"
              : "#VALUE!",
            error instanceof Error
              ? error.message
              : "Invalid reusable formula.",
          );
        }
        tokenCache.set(sourceText, cachedTokens);
      }
      return cachedTokens;
    },
    resolveNameBinding = (reference, fallbackSheetId) => {
      const targetSheetId = reference.qualifier
        ? workbook.sheets.find(
            (sheet) =>
              sheet.name.toLowerCase() === reference.qualifier.toLowerCase(),
          )?.id
        : fallbackSheetId;
      if (!targetSheetId) return;
      const bindingName = normalizeBindingName(reference.name),
        found =
          workbook.definedNames?.find(
            (definedName) =>
              definedName.scopeSheetId === targetSheetId &&
              definedName.name.toLowerCase() === bindingName,
          ) ??
          workbook.definedNames?.find(
            (definedName) =>
              !definedName.scopeSheetId &&
              definedName.name.toLowerCase() === bindingName,
          );
      return found
        ? { found: found, sheetId: found.scopeSheetId ?? targetSheetId }
        : undefined;
    },
    parseOrIssue = (text, contextSheetId, externalValues) => {
      try {
        return evaluationContext.parse(text, contextSheetId, externalValues);
      } catch (error) {
        const errorDetails =
          error && typeof error == "object" && "details" in error
            ? error.details
            : undefined;
        return evaluationContext.isIssue(errorDetails)
          ? errorDetails
          : evaluationContext.isIssue(error)
            ? error
            : evaluationContext.issue("#ERROR!");
      }
    },
    resolveSymbol = (
      symbolNode,
      symbolSource,
      symbolBindings,
      symbolSheetId,
      symbolDepth,
    ) => {
      const bindingName = normalizeBindingName(symbolNode.name);
      if (!symbolNode.qualifier && symbolBindings.has(bindingName))
        return symbolBindings.get(bindingName);
      if (!symbolNode.qualifier && /^(?:TRUE|FALSE)$/i.test(symbolNode.name))
        return bindingName === "true";
      const resolvedName = resolveNameBinding(symbolNode, symbolSheetId);
      if (!resolvedName)
        return raiseCalcError(
          "#NAME?",
          `Unknown name or unsupported custom function: ${symbolSource.slice(symbolNode.start, symbolNode.end)}.`,
        );
      const nameKey = `${resolvedName.sheetId}:${resolvedName.found.name.toLowerCase()}`;
      if (
        evaluationContext.allowName &&
        !evaluationContext.allowName(resolvedName.sheetId)
      )
        return raiseCalcError(
          "#BLOCKED!",
          "The defined name is outside the authorized inspection context.",
        );
      if (activeNameKeys.has(nameKey))
        return raiseCalcError(
          "#CYCLE!",
          `Circular defined name: ${resolvedName.found.name}.`,
        );
      activeNameKeys.add(nameKey);
      try {
        return evaluateTokens(
          cachedTokenize(resolvedName.found.refersTo),
          resolvedName.found.refersTo,
          new Map(),
          resolvedName.sheetId,
          symbolDepth + 1,
        );
      } finally {
        activeNameKeys.delete(nameKey);
      }
    },
    callLambda = (closure, args, lambdaDepth) => {
      if (!isFormulaClosure(closure))
        return raiseCalcError(
          "#VALUE!",
          "Only a LAMBDA formula can be called with arguments.",
        );
      if (args.length !== closure.parameters.length)
        return raiseCalcError(
          "#VALUE!",
          `LAMBDA expects ${closure.parameters.length} argument(s); received ${args.length}.`,
        );
      if (++callCount > reusableFormulaLimits.calls)
        return raiseCalcError(
          "#LIMIT!",
          "Reusable formula evaluation exceeds 1,024 calls.",
        );
      const closureEnvironment = new Map(closure.environment);
      return (
        closure.parameters.forEach((parameter, index) =>
          closureEnvironment.set(parameter, retainValues(args[index])),
        ),
        evaluateTokens(
          closure.body,
          closure.source,
          closureEnvironment,
          closure.sheetId,
          lambdaDepth + 1,
        )
      );
    },
    evaluateNode = (node, source, localBindings, sheetId, callDepth) => {
      if ((chargeWork(1, callDepth), node.kind === "symbol"))
        return resolveSymbol(node, source, localBindings, sheetId, callDepth);
      if (node.kind === "group")
        return evaluateTokens(
          node.body,
          source,
          localBindings,
          sheetId,
          callDepth + 1,
        );
      if (node.kind === "text") return parseOrIssue(node.text, sheetId, []);
      const localBindingFunction = validateLetLambdaCall(node);
      if (localBindingFunction === "LET") {
        const letBindings = new Map(localBindings);
        for (let argIndex = 0; argIndex < node.args.length - 1; argIndex += 2) {
          const bindingValue = evaluateTokens(
              node.args[argIndex + 1],
              source,
              letBindings,
              sheetId,
              callDepth + 1,
            ),
            bindingName = normalizeBindingName(
              validateBindingName(node.args[argIndex]).name,
            );
          (letBindings.set(bindingName, retainValues(bindingValue)),
            evaluationContext.step?.(
              "binding",
              `LET ${bindingName}`,
              bindingValue,
            ));
        }
        return evaluateTokens(
          node.args.at(-1),
          source,
          letBindings,
          sheetId,
          callDepth + 1,
        );
      }
      if (localBindingFunction === "LAMBDA")
        return {
          kind: "formula-closure",
          parameters: node.args
            .slice(0, -1)
            .map((parameterNode) =>
              normalizeBindingName(
                validateBindingName(parameterNode, true).name,
              ),
            ),
          body: node.args.at(-1),
          source: source,
          environment: new Map(localBindings),
          sheetId: sheetId,
        };
      const functionName =
          node.callee.kind === "symbol" && !node.callee.qualifier
            ? normalizeFunctionName(node.callee.name)
            : "",
        isLocalCallee =
          node.callee.kind === "symbol" &&
          !node.callee.qualifier &&
          localBindings.has(normalizeBindingName(node.callee.name));
      if (
        !isLocalCallee &&
        functionName === "IF" &&
        node.args.length >= 2 &&
        node.args.length <= 3
      ) {
        const conditionValue = evaluateTokens(
          node.args[0],
          source,
          localBindings,
          sheetId,
          callDepth + 1,
        );
        if (evaluationContext.isIssue(conditionValue)) return conditionValue;
        if (!Array.isArray(conditionValue)) {
          if (
            typeof conditionValue == "string" &&
            !/^(?:TRUE|FALSE)$/i.test(conditionValue)
          )
            return evaluationContext.issue("#VALUE!");
          const isConditionTrue =
            typeof conditionValue == "string"
              ? conditionValue.toUpperCase() === "TRUE"
              : !!conditionValue;
          return (
            evaluationContext.step?.(
              "branch",
              `IF: ${isConditionTrue ? "true" : "false"} branch selected`,
              conditionValue,
            ),
            node.args[isConditionTrue ? 1 : 2]
              ? evaluateTokens(
                  node.args[isConditionTrue ? 1 : 2],
                  source,
                  localBindings,
                  sheetId,
                  callDepth + 1,
                )
              : false
          );
        }
      }
      if (
        !isLocalCallee &&
        /^(?:IFERROR|IFNA)$/.test(functionName) &&
        node.args.length === 2
      ) {
        const guardedValue = evaluateTokens(
          node.args[0],
          source,
          localBindings,
          sheetId,
          callDepth + 1,
        );
        if (!Array.isArray(guardedValue)) {
          if (
            evaluationContext.isIssue(guardedValue) &&
            ["#LIMIT!", "#BLOCKED!", "#CYCLE!", "#ERROR!"].includes(
              guardedValue.error,
            )
          )
            return guardedValue;
          const shouldUseFallback =
            evaluationContext.isIssue(guardedValue) &&
            (functionName === "IFERROR" || guardedValue.error === "#N/A");
          return (
            evaluationContext.step?.(
              "branch",
              `${functionName}: ${shouldUseFallback ? "fallback" : "value"} selected`,
              guardedValue,
            ),
            shouldUseFallback
              ? evaluateTokens(
                  node.args[1],
                  source,
                  localBindings,
                  sheetId,
                  callDepth + 1,
                )
              : guardedValue
          );
        }
      }
      if (
        !isLocalCallee &&
        functionName &&
        evaluationContext.supported.has(functionName)
      ) {
        const supportedArgValues = [],
          argTexts = node.args.map((arg) =>
            serializeTokens(
              arg,
              source,
              localBindings,
              sheetId,
              callDepth + 1,
              supportedArgValues,
            ),
          );
        return parseOrIssue(
          `${functionName}(${argTexts.join(",")})`,
          sheetId,
          supportedArgValues,
        );
      }
      const calleeValue = evaluateNode(
          node.callee,
          source,
          localBindings,
          sheetId,
          callDepth + 1,
        ),
        argumentValues = node.args.map((arg) =>
          dropBlankTextTokens(arg).length
            ? evaluateTokens(arg, source, localBindings, sheetId, callDepth + 1)
            : null,
        ),
        callResult = callLambda(calleeValue, argumentValues, callDepth);
      return (
        evaluationContext.step?.(
          "function",
          source.slice(node.callee.start, node.callee.end),
          callResult,
        ),
        callResult
      );
    },
    serializeTokens = (
      tokenList,
      tokenSource,
      tokenBindings,
      tokenSheetId,
      tokenDepth,
      pushedValues,
    ) =>
      tokenList
        .map((token) => {
          if (token.kind === "text") {
            const scientificMatch =
              /^(\d+(?:\.\d*)?|\.\d+)[eE]([+-]?\d+)$/.exec(token.text);
            return scientificMatch
              ? `(${scientificMatch[1]}*10^(${Number(scientificMatch[2])}))`
              : token.text;
          }
          const tokenValue = evaluateNode(
            token,
            tokenSource,
            tokenBindings,
            tokenSheetId,
            tokenDepth,
          );
          return isFormulaClosure(tokenValue)
            ? raiseCalcError(
                "#CALC!",
                "A LAMBDA must be called with arguments or saved as a reusable name.",
              )
            : (retainValues(tokenValue),
              pushedValues.push(
                Array.isArray(tokenValue)
                  ? tokenValue.map((valueRow) =>
                      Array.isArray(valueRow) ? [...valueRow] : valueRow,
                    )
                  : tokenValue,
              ),
              `${y5}(${pushedValues.length - 1})`);
        })
        .join(""),
    evaluateTokens = (
      tokenNodes,
      evaluationSource,
      scopeBindings,
      formulaSheetId,
      evaluationDepth,
    ) => {
      try {
        chargeWork(tokenNodes.length, evaluationDepth);
        const meaningfulTokens = dropBlankTextTokens(tokenNodes);
        if (
          meaningfulTokens.length === 1 &&
          meaningfulTokens[0].kind !== "text"
        )
          return evaluateNode(
            meaningfulTokens[0],
            evaluationSource,
            scopeBindings,
            formulaSheetId,
            evaluationDepth,
          );
        const collectedValues = [];
        return parseOrIssue(
          serializeTokens(
            tokenNodes,
            evaluationSource,
            scopeBindings,
            formulaSheetId,
            evaluationDepth,
            collectedValues,
          ),
          formulaSheetId,
          collectedValues,
        );
      } catch (error) {
        if (
          evaluationContext.isIssue(error) &&
          !["#LIMIT!", "#BLOCKED!", "#CYCLE!", "#ERROR!"].includes(error.error)
        )
          return error;
        throw error;
      }
    };
  try {
    const result = evaluateTokens(
      cachedTokenize(formulaText),
      formulaText,
      new Map(),
      sheetId,
      0,
    );
    return isFormulaClosure(result)
      ? raiseCalcError(
          "#CALC!",
          "A LAMBDA must be called with arguments or saved as a reusable name.",
        )
      : result;
  } catch (error) {
    if (evaluationContext.isIssue(error)) return error;
    throw error;
  }
}
function summarizeWorksheets(workbook, options = {}) {
  return workbook.sheets
    .filter(
      (sheet) =>
        !options.allowedSheetIds || options.allowedSheetIds.includes(sheet.id),
    )
    .map((sheet) => ({
      sheetId: sheet.id,
      name: sheet.name,
      visibility: sheet.veryHidden
        ? "veryHidden"
        : sheet.hidden
          ? "hidden"
          : "visible",
      tabColor: sheet.tabColor ?? null,
      storedCellCount: Object.keys(sheet.cells).length,
      formulaCellCount: Object.values(sheet.cells).filter(
        (cell) =>
          !cell.literal &&
          typeof cell.raw == "string" &&
          cell.raw.startsWith("="),
      ).length,
      calculationParticipation: "included-when-referenced",
      structureLocked: !!workbook.structureProtection?.enabled,
    }));
}
const createUuid = () => crypto.randomUUID();

const isPlainObject = (value) =>
  !!value &&
  typeof value == "object" &&
  !Array.isArray(value) &&
  [null, Object.prototype].includes(Object.getPrototypeOf(value));
function validateWorksheetName(sheetName) {
  return (
    (typeof sheetName != "string" ||
      !sheetName.trim() ||
      sheetName.length > 31 ||
      /[\\/\[\]*?:\u0000-\u001f]/.test(sheetName) ||
      /^'|'$/.test(sheetName)) &&
      throwWorkbookError(
        "Sheet names must contain 1–31 characters and cannot contain control characters, Excel-reserved characters, or begin or end with an apostrophe.",
      ),
    sheetName
  );
}
function assertUniqueWorksheetName(workbook, sheetName, excludedSheetId) {
  const validatedName = validateWorksheetName(sheetName);
  return (
    workbook.sheets.some(
      (sheet) =>
        sheet.id !== excludedSheetId &&
        sheet.name.toLowerCase() === validatedName.toLowerCase(),
    ) && throwWorkbookError("A worksheet with that name already exists."),
    validatedName
  );
}
const findWorksheetById = (workbook, sheetId) =>
    workbook.sheets.find((sheet) => sheet.id === sheetId) ??
    throwWorkbookError("The worksheet was not found."),
  isVisibleWorksheet = (sheet) => !sheet.hidden && !sheet.veryHidden;
function validateTabColor(tabColor) {
  return (
    (typeof tabColor != "string" || !/^#[a-f\d]{6}$/i.test(tabColor)) &&
      throwWorkbookError("Tab colors must use a six-digit hexadecimal color."),
    tabColor.toLowerCase()
  );
}
function validateWorksheetVisibility(sheet, targetSheet) {
  (validateWorksheetName(targetSheet.name),
    sheet.veryHidden !== undefined &&
      (typeof sheet.veryHidden != "boolean" &&
        throwWorkbookError("The very-hidden flag must be boolean."),
      sheet.veryHidden &&
        sheet.hidden === false &&
        throwWorkbookError("A very-hidden worksheet must also be hidden."),
      (targetSheet.veryHidden = sheet.veryHidden),
      sheet.veryHidden && (targetSheet.hidden = true)),
    sheet.tabColor !== undefined &&
      (targetSheet.tabColor = validateTabColor(sheet.tabColor)));
}
function validateWorkbookStructureProtection(structureProtection, workbook) {
  const protection = structureProtection;
  if (protection !== undefined) {
    (!isPlainObject(protection) ||
      protection.enabled !== true ||
      Object.keys(protection).some(
        (key) => !["enabled", "passwordHash", "salt"].includes(key),
      )) &&
      throwWorkbookError("Invalid workbook structure protection.");
    const hasPasswordHash = protection.passwordHash !== undefined,
      hasSalt = protection.salt !== undefined;
    ((hasPasswordHash !== hasSalt ||
      (hasPasswordHash &&
        (typeof protection.passwordHash != "string" ||
          !/^[a-f\d]{64}$/i.test(protection.passwordHash) ||
          typeof protection.salt != "string" ||
          protection.salt.length < 16 ||
          protection.salt.length > 100))) &&
      throwWorkbookError(
        "Workbook structure protection requires a valid salted password hash.",
      ),
      (workbook.structureProtection = {
        enabled: true,
        ...(hasPasswordHash
          ? {
              passwordHash: String(protection.passwordHash).toLowerCase(),
              salt: String(protection.salt),
            }
          : {}),
      }));
  }
  workbook.sheets.some(isVisibleWorksheet) ||
    throwWorkbookError("At least one worksheet must remain visible.");
}
function assertWorkbookStructureEditable(workbook) {
  workbook.structureProtection?.enabled &&
    throwWorkbookError(
      "Workbook structure is protected. Unprotect the structure before changing worksheets.",
    );
}
const S5 = String.raw`(?:'(?:[^']|'')+'|[^\s+\-*/^&=<>(),;:'"\[\]!]+)`,
  sheetPrefixPattern = new RegExp(`^(${S5})!`),
  Bne = new RegExp(`^(${S5}):(${S5})!`),
  unquoteSheetName = (quotedSheetName) =>
    quotedSheetName.startsWith("'")
      ? quotedSheetName.slice(1, -1).replace(/''/g, "'")
      : quotedSheetName,
  quoteSheetName = (sheetName) => `'${sheetName.replace(/'/g, "''")}'`;
function findClosingBracketIndex(formulaText, startIndex) {
  let scanIndex = startIndex,
    bracketDepth = 0;
  do
    (formulaText[scanIndex] === "[" && bracketDepth++,
      formulaText[scanIndex] === "]" && bracketDepth--,
      scanIndex++);
  while (scanIndex < formulaText.length && bracketDepth > 0);
  return bracketDepth === 0 ? scanIndex : -1;
}
function Q6(formulaText, end) {
  const matchLength = Math.max(
    0,
    ...[
      /^\$?[A-Za-z]{1,3}\$?[1-9]\d*(?::\$?[A-Za-z]{1,3}\$?[1-9]\d*)?/,
      /^\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}/,
      /^\$?[1-9]\d*:\$?[1-9]\d*/,
      /^[A-Za-z_\\][A-Za-z0-9_.\\]*/,
    ].map((pattern) => pattern.exec(formulaText.slice(end))?.[0].length ?? 0),
  );
  if (!matchLength) return end;
  let referenceEnd = end + matchLength;
  return (
    formulaText[referenceEnd] === "[" &&
      (referenceEnd = findClosingBracketIndex(formulaText, referenceEnd)),
    referenceEnd < end ||
    /[\w.$]/.test(formulaText[referenceEnd] ?? "") ||
    /^\s*\(/.test(formulaText.slice(referenceEnd))
      ? end
      : (formulaText[referenceEnd] === "#" && referenceEnd++, referenceEnd)
  );
}
function parseWorksheetQualifiedReference(formulaText) {
  const references = [];
  let end = 0;
  for (; end < formulaText.length;) {
    const tokenStart = end,
      currentChar = formulaText[end];
    if (currentChar === '"') {
      for (end++; end < formulaText.length;)
        if (formulaText[end++] === '"')
          if (formulaText[end] === '"') end++;
          else break;
      continue;
    }
    const isTokenStart = end === 0 || !/[\w.$]/.test(formulaText[end - 1]);
    isTokenStart &&
      /^(?:_xlfn\.)?INDIRECT\s*\(/i.test(formulaText.slice(end)) &&
      references.push({
        start: tokenStart,
        end: tokenStart,
        text: "INDIRECT",
        unsupported: "Dynamic INDIRECT references cannot be safely repaired.",
      });
    const sheetRangeMatch = isTokenStart
      ? Bne.exec(formulaText.slice(end))
      : null;
    if (sheetRangeMatch) {
      end += sheetRangeMatch[0].length;
      const referenceEnd = Q6(formulaText, end);
      ((end = Math.max(end, referenceEnd)),
        references.push({
          start: tokenStart,
          end: end,
          text: formulaText.slice(tokenStart, end),
          sheets: [
            unquoteSheetName(sheetRangeMatch[1]),
            unquoteSheetName(sheetRangeMatch[2]),
          ],
          unsupported:
            "Three-dimensional worksheet references require explicit references before this operation.",
        }));
      continue;
    }
    const sheetPrefixMatch = isTokenStart
      ? sheetPrefixPattern.exec(formulaText.slice(end))
      : null;
    if (sheetPrefixMatch) {
      const sheet = unquoteSheetName(sheetPrefixMatch[1]);
      end += sheetPrefixMatch[0].length;
      const sheetReferenceEnd = Q6(formulaText, end),
        reference = formulaText.slice(end, sheetReferenceEnd);
      if (
        ((end = Math.max(end, sheetReferenceEnd)),
        sheet.includes("[") || sheet.includes("]"))
      )
        continue;
      references.push({
        start: tokenStart,
        end: end,
        text: formulaText.slice(tokenStart, end),
        ...(sheet.includes(":")
          ? {
              sheets: sheet.split(":"),
              unsupported:
                "Three-dimensional worksheet references require explicit references before this operation.",
            }
          : {
              sheet: sheet,
              reference: reference,
              ...(reference
                ? {}
                : {
                    unsupported:
                      "This worksheet-qualified expression is not a supported reference.",
                  }),
            }),
      });
      continue;
    }
    const tableMatch = isTokenStart
      ? /^([A-Za-z_][A-Za-z0-9_.]*)\[/.exec(formulaText.slice(end))
      : null;
    if (tableMatch) {
      const closingBracketIndex = findClosingBracketIndex(
        formulaText,
        end + tableMatch[1].length,
      );
      if (closingBracketIndex > 0) {
        ((end = closingBracketIndex),
          references.push({
            start: tokenStart,
            end: closingBracketIndex,
            text: formulaText.slice(tokenStart, closingBracketIndex),
            table: tableMatch[1],
          }));
        continue;
      }
    }
    if (currentChar === "[") {
      const closingBracketIndex = findClosingBracketIndex(formulaText, end);
      if (closingBracketIndex > 0) {
        end = closingBracketIndex;
        const trailingSheetMatch = sheetPrefixPattern.exec(
          formulaText.slice(end),
        );
        trailingSheetMatch &&
          ((end += trailingSheetMatch[0].length),
          (end = Math.max(end, Q6(formulaText, end))));
        continue;
      }
    }
    if (currentChar === "'") {
      for (end++; end < formulaText.length;)
        if (formulaText[end++] === "'")
          if (formulaText[end] === "'") end++;
          else break;
      continue;
    }
    end++;
  }
  return references;
}
function xP(workbook, workbookModel, worksheetById) {
  if (workbook.sheet)
    return workbook.sheet.toLowerCase() === worksheetById.name.toLowerCase();
  if (workbook.sheets) {
    const sheetIndexes = workbook.sheets.map((sheet) =>
        workbookModel.sheets.findIndex(
          (modelSheet) => modelSheet.name.toLowerCase() === sheet.toLowerCase(),
        ),
      ),
      worksheetIndex = workbookModel.sheets.findIndex(
        (sheet) => sheet.id === worksheetById.id,
      );
    return (
      sheetIndexes.some((sheetIndex) => sheetIndex < 0) ||
      (worksheetIndex >= Math.min(...sheetIndexes) &&
        worksheetIndex <= Math.max(...sheetIndexes))
    );
  }
  return !!workbook.unsupported;
}
function rewriteStructuredTableReference(
  formulaText,
  workbook,
  worksheetById,
  newSheetName,
  shouldRewriteTables = false,
) {
  const localTableNames = new Set(
    (worksheetById.tables ?? []).map((table) => table.name.toLowerCase()),
  );
  let rewrittenText = "",
    copiedUntil = 0;
  for (const qualifiedReference of parseWorksheetQualifiedReference(
    formulaText,
  )) {
    const isTargetSheetReference = xP(
        qualifiedReference,
        workbook,
        worksheetById,
      ),
      isOwnTableReference =
        shouldRewriteTables &&
        qualifiedReference.table &&
        localTableNames.has(qualifiedReference.table.toLowerCase());
    (qualifiedReference.unsupported &&
      isTargetSheetReference &&
      throwWorkbookError(qualifiedReference.unsupported),
      !(!isTargetSheetReference && !isOwnTableReference) &&
        ((rewrittenText += formulaText.slice(
          copiedUntil,
          qualifiedReference.start,
        )),
        (rewrittenText +=
          newSheetName !== undefined
            ? `${quoteSheetName(newSheetName)}!${qualifiedReference.reference}`
            : "#REF!"),
        (copiedUntil = qualifiedReference.end)));
  }
  return rewrittenText + formulaText.slice(copiedUntil);
}
function visitTableFormulasOutsideSheet(
  workbook,
  visitFormula,
  excludedSheetId,
) {
  for (const sheet of workbook.sheets)
    if (sheet.id !== excludedSheetId) {
      for (const table of sheet.tables ?? [])
        for (const column of table.columns)
          column.calculatedFormula &&
            visitFormula(column.calculatedFormula, sheet.id);
      for (const rule of [
        ...(sheet.conditionalFormats ?? []),
        ...(sheet.validations ?? []),
      ])
        (rule.kind === "formula" || rule.kind === "custom") &&
          rule.formula &&
          visitFormula(rule.formula, sheet.id, rule.range.split(":")[0]);
      for (const [address, cell] of Object.entries(sheet.cells))
        !cell.literal &&
          typeof cell.raw == "string" &&
          cell.raw.startsWith("=") &&
          visitFormula(cell.raw, sheet.id, address);
    }
  for (const definedName of workbook.definedNames ?? [])
    visitFormula(
      definedName.refersTo,
      definedName.scopeSheetId ?? "",
      undefined,
    );
}
function createUniqueWorksheetName(workbook, baseName) {
  for (
    let suffixNumber = 2;
    suffixNumber < Number.MAX_SAFE_INTEGER;
    suffixNumber++
  ) {
    const suffix = ` (${suffixNumber})`,
      candidateName = baseName.slice(0, 31 - suffix.length) + suffix;
    if (
      !workbook.sheets.some(
        (sheet) => sheet.name.toLowerCase() === candidateName.toLowerCase(),
      )
    )
      return candidateName;
  }
  return throwWorkbookError("No unique worksheet name is available.");
}
function describeSheetDeletionImpacts(workbook, sheetId) {
  const worksheetById = findWorksheetById(workbook, sheetId),
    impacts = [],
    deletedTableNames = new Set(
      (worksheetById.tables ?? []).map((deletedTable) =>
        deletedTable.name.toLowerCase(),
      ),
    ),
    addReferenceImpact = (formulaText, impact) => {
      const worksheetQualifiedReference =
          parseWorksheetQualifiedReference(formulaText),
        affectedReferences = worksheetQualifiedReference.filter(
          (qualifiedReference) =>
            xP(qualifiedReference, workbook, worksheetById) ||
            (qualifiedReference.table &&
              deletedTableNames.has(qualifiedReference.table.toLowerCase())),
        );
      affectedReferences.some(
        (affectedReference) => affectedReference.unsupported,
      )
        ? impacts.push({
            ...impact,
            kind: "unsupported-reference",
            label: `${impact.label}: ${affectedReferences.find((unsupportedReference) => unsupportedReference.unsupported).unsupported}`,
          })
        : affectedReferences.length && impacts.push(impact);
    };
  for (const sheet of workbook.sheets)
    if (sheet.id !== sheetId) {
      for (const sheetTable of sheet.tables ?? [])
        for (const column of sheetTable.columns)
          column.calculatedFormula &&
            addReferenceImpact(column.calculatedFormula, {
              kind: "table-column",
              label: `${sheet.name}: ${sheetTable.name}[${column.name}] calculated rule will contain #REF!`,
              sheetId: sheet.id,
            });
      for (const [address, cell] of Object.entries(sheet.cells))
        !cell.literal &&
          typeof cell.raw == "string" &&
          cell.raw.startsWith("=") &&
          addReferenceImpact(cell.raw, {
            kind: "formula",
            label: `${sheet.name}!${address} will contain #REF!`,
            sheetId: sheet.id,
            address: address,
          });
      for (const rule of [
        ...(sheet.conditionalFormats ?? []),
        ...(sheet.validations ?? []),
      ])
        (rule.kind === "formula" || rule.kind === "custom") &&
          rule.formula &&
          addReferenceImpact(rule.formula, {
            kind: "formula",
            label: `${sheet.name}: ${rule.kind === "custom" ? "validation" : "conditional"} rule ${rule.id} will contain #REF!`,
            sheetId: sheet.id,
            address: rule.range.split(":")[0],
          });
      for (const pivot of sheet.pivots ?? [])
        pivot.source.sheetId === sheetId &&
          impacts.push({
            kind: "pivot-refresh",
            label: `${sheet.name}: ${pivot.name} refresh settings will be removed; output cells retained`,
            sheetId: sheet.id,
          });
      for (const query of sheet.queries ?? [])
        query.source.sheetId === sheetId &&
          impacts.push({
            kind: "query-refresh",
            label: `${sheet.name}: ${query.name} refresh settings will be removed; output cells retained`,
            sheetId: sheet.id,
          });
    }
  for (const definedName of workbook.definedNames ?? [])
    definedName.scopeSheetId === sheetId
      ? impacts.push({
          kind: "local-name",
          label: `${worksheetById.name}!${definedName.name} local name will be removed`,
          sheetId: sheetId,
        })
      : addReferenceImpact(definedName.refersTo, {
          kind: "defined-name",
          label: `${definedName.name} will contain #REF!`,
          ...(definedName.scopeSheetId
            ? { sheetId: definedName.scopeSheetId }
            : {}),
        });
  for (const formulaWatch of workbook.formulaWatches ?? [])
    !formulaWatch.orphaned &&
      formulaWatch.anchor.sheetId === sheetId &&
      impacts.push({
        kind: "formula-watch",
        label: `${worksheetById.name}!${formulaWatch.anchor.address} watch will be retained as a deleted-sheet orphan`,
        sheetId: sheetId,
        address: formulaWatch.anchor.address,
      });
  return impacts;
}
function validateWorkbookPassword(password) {
  return (
    (typeof password != "string" || password.length > 200) &&
      throwWorkbookError("Passwords must be text of at most 200 characters."),
    password
  );
}
async function hashWorkbookStructurePassword(password, salt) {
  const digest = await crypto.subtle.digest(
    "SHA-256",
    new TextEncoder().encode(`workbook-structure\0${salt}\0${password}`),
  );
  return [...new Uint8Array(digest)]
    .map((byteValue) => byteValue.toString(16).padStart(2, "0"))
    .join("");
}
const tableColumnDataTypes = ["any", "text", "number", "boolean", "date"];
function validateTableColumnType(dataType) {
  if (!tableColumnDataTypes.includes(dataType))
    throw Error("Choose a supported table column type.");
  return dataType;
}
function applyTableSourceEvidence(workbook, incomingWorkbook) {
  if (!workbook.tableEvidence?.length) return incomingWorkbook;
  const existingEvidenceById = new Map(
    workbook.tableEvidence.map((evidenceItem) => [
      evidenceItem.id,
      evidenceItem,
    ]),
  );
  for (const incomingEvidence of incomingWorkbook.tableEvidence ?? []) {
    const existingEvidence = existingEvidenceById.get(incomingEvidence.id);
    if (
      existingEvidence &&
      JSON.stringify(existingEvidence) !== JSON.stringify(incomingEvidence)
    )
      throw Error("Saved table source evidence cannot be changed.");
  }
  return {
    ...incomingWorkbook,
    tableEvidence: [
      ...workbook.tableEvidence,
      ...(incomingWorkbook.tableEvidence ?? []).filter(
        (evidence) => !existingEvidenceById.has(evidence.id),
      ),
    ],
  };
}
const isNullOrEmptyValue = (value) => value === null || value === "";
function parseIsoDateText(dateText) {
  if (
    typeof dateText != "string" ||
    !/^\d{4}-\d{2}-\d{2}$/.test(dateText) ||
    dateText.startsWith("0000")
  )
    return false;
  const timestamp = Date.parse(dateText + "T00:00:00.000Z");
  return (
    Number.isFinite(timestamp) &&
    new Date(timestamp).toISOString().slice(0, 10) === dateText
  );
}
function coerceImportedValue(importedValue, dataType = "any") {
  return (
    dataType === "any" ||
    isNullOrEmptyValue(importedValue) ||
    (dataType === "date"
      ? parseIsoDateText(importedValue)
      : typeof importedValue === (dataType === "text" ? "string" : dataType) &&
        (dataType !== "number" || Number.isFinite(importedValue)))
  );
}
function convertTextForDataType(importedValue, options) {
  if (coerceImportedValue(importedValue, options.dataType))
    return importedValue;
  if (options.dataType === "text") return String(importedValue);
  if (typeof importedValue != "string")
    throw Error("This value cannot be converted without guessing.");
  const trimmedText = importedValue.trim();
  if (options.dataType === "boolean" && /^(true|false)$/i.test(trimmedText))
    return /^true$/i.test(trimmedText);
  if (options.dataType === "date" && parseIsoDateText(trimmedText))
    return trimmedText;
  if (options.dataType === "number") {
    const decimalSeparator = options.decimal ?? ".",
      groupSeparator = options.group ?? ",",
      escapeRegExp = (text) => text.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"),
      integerPattern = groupSeparator
        ? `(?:\\d+|\\d{1,3}(?:${escapeRegExp(groupSeparator)}\\d{3})+)`
        : "\\d+";
    if (
      !new RegExp(
        `^[+-]?(?:${integerPattern}(?:${escapeRegExp(decimalSeparator)}\\d+)?|${escapeRegExp(decimalSeparator)}\\d+)$`,
      ).test(trimmedText)
    )
      throw Error(
        "Use a number with the selected decimal/group separators; currency, percentages and scientific text require a separate cleanup.",
      );
    const normalizedNumberText = (
        groupSeparator
          ? trimmedText.split(groupSeparator).join("")
          : trimmedText
      ).replace(decimalSeparator, "."),
      numericValue = Number(normalizedNumberText),
      significantDigits = normalizedNumberText
        .replace(/^[+-]?0*/, "")
        .replace(".", "")
        .replace(/0+$/, "");
    if (
      !Number.isFinite(numericValue) ||
      significantDigits.length > 15 ||
      (Number.isInteger(numericValue) && !Number.isSafeInteger(numericValue))
    )
      throw Error(
        "Conversion would exceed safe numeric precision. Keep the original text.",
      );
    return numericValue;
  }
  throw Error(
    options.dataType === "date"
      ? "Dates must be valid YYYY-MM-DD text; no locale or date-system inference."
      : "Boolean text must be TRUE or FALSE.",
  );
}
const supportedChartKinds = [
  "column",
  "bar",
  "line",
  "area",
  "pie",
  "donut",
  "scatter",
  "combo",
  "waterfall",
  "histogram",
  "sparkline",
];

function validateChartSettings(chartOptions) {
  if (chartOptions === undefined) return {};
  (!isNonArrayObject(chartOptions) ||
    Object.keys(chartOptions).some(
      (chartOption) =>
        ![
          "altText",
          "xTitle",
          "yTitle",
          "units",
          "yScale",
          "xScale",
          "showLegend",
          "series",
          "includedRows",
          "histogramBins",
          "sparklineKind",
        ].includes(chartOption),
    )) &&
    failWithMessage("Invalid or unknown chart settings.");
  const validatedSettings = {};
  for (const textField of ["altText", "xTitle", "yTitle", "units"])
    chartOptions[textField] !== undefined &&
      ((typeof chartOptions[textField] != "string" ||
        chartOptions[textField].length >
          (textField === "altText" ? 2e3 : 200)) &&
        failWithMessage(`Invalid chart ${textField}.`),
      (validatedSettings[textField] = chartOptions[textField]));
  for (const scaleField of ["xScale", "yScale"])
    chartOptions[scaleField] !== undefined &&
      (["linear", "log"].includes(String(chartOptions[scaleField])) ||
        failWithMessage("Chart scales must be linear or log."),
      (validatedSettings[scaleField] = chartOptions[scaleField]));
  if (
    (chartOptions.showLegend !== undefined &&
      (typeof chartOptions.showLegend != "boolean" &&
        failWithMessage("Invalid chart legend setting."),
      (validatedSettings.showLegend = chartOptions.showLegend)),
    chartOptions.histogramBins !== undefined &&
      ((!Number.isInteger(chartOptions.histogramBins) ||
        Number(chartOptions.histogramBins) < 1 ||
        Number(chartOptions.histogramBins) > 50) &&
        failWithMessage("Choose 1 to 50 histogram bins."),
      (validatedSettings.histogramBins = Number(chartOptions.histogramBins))),
    chartOptions.sparklineKind !== undefined &&
      (["line", "column", "winLoss"].includes(
        String(chartOptions.sparklineKind),
      ) || failWithMessage("Invalid sparkline style."),
      (validatedSettings.sparklineKind = chartOptions.sparklineKind)),
    chartOptions.series !== undefined)
  ) {
    (!Array.isArray(chartOptions.series) || chartOptions.series.length > 20) &&
      failWithMessage("Select at most 20 chart series.");
    const seenColumns = new Set();
    validatedSettings.series = chartOptions.series.map((sery) => {
      ((!isNonArrayObject(sery) ||
        Object.keys(sery).some(
          (key) => !["column", "name", "kind"].includes(key),
        ) ||
        !Number.isInteger(sery.column) ||
        Number(sery.column) < 1 ||
        Number(sery.column) > 16384 ||
        seenColumns.has(Number(sery.column))) &&
        failWithMessage(
          "Chart series require unique one-based worksheet columns.",
        ),
        seenColumns.add(Number(sery.column)));
      const normalizedSeries = { column: Number(sery.column) };
      return (
        sery.name !== undefined &&
          ((typeof sery.name != "string" ||
            !sery.name.trim() ||
            sery.name.length > 200) &&
            failWithMessage("Chart series names require 1 to 200 characters."),
          (normalizedSeries.name = sery.name)),
        sery.kind !== undefined &&
          (["column", "line"].includes(String(sery.kind)) ||
            failWithMessage("Combo series must be column or line."),
          (normalizedSeries.kind = sery.kind)),
        normalizedSeries
      );
    });
  }
  return (
    chartOptions.includedRows !== undefined &&
      ((!Array.isArray(chartOptions.includedRows) ||
        chartOptions.includedRows.length > 500 ||
        chartOptions.includedRows.some(
          (includedRow) =>
            !Number.isInteger(includedRow) ||
            includedRow < 1 ||
            includedRow > 1048576,
        ) ||
        new Set(chartOptions.includedRows).size !==
          chartOptions.includedRows.length) &&
        failWithMessage(
          "Chart filters require at most 500 unique one-based worksheet rows.",
        ),
      (validatedSettings.includedRows = [...chartOptions.includedRows].sort(
        (leftRow, rightRow) => leftRow - rightRow,
      ))),
    validatedSettings
  );
}
function formatDiagnosticsMessage(validationResult) {
  const errorDiagnostics = validationResult.diagnostics.filter(
    (diagnostic) => diagnostic.severity === "error",
  );
  errorDiagnostics.length &&
    failWithMessage(
      errorDiagnostics
        .slice(0, 5)
        .map((errorDiagnostic) => errorDiagnostic.message)
        .join(" "),
    );
}
const Kne = 100,
  maxPictureHashBytes = 2 * 1024 * 1024,
  Yne = 500,
  validateDrawingNumber = (value, minimum, maximum, fieldName) => {
    if (
      typeof value != "number" ||
      !Number.isFinite(value) ||
      value < minimum ||
      value > maximum
    )
      throw new Error(`Invalid drawing ${fieldName}.`);
    return value;
  },
  validateDrawingText = (value, maxLength, fieldName) => {
    if (typeof value != "string" || value.length > maxLength)
      throw new Error(`Invalid drawing ${fieldName}.`);
    return value;
  },
  validateDrawingColor = (value) => {
    if (typeof value != "string" || !/^#[\da-f]{6}$/i.test(value))
      throw new Error("Drawing colors must be six-digit hex colors.");
    return value;
  };
function hashPictureBytes(bytes) {
  if (!(bytes instanceof Uint8Array) || bytes.length > maxPictureHashBytes)
    throw Error("Picture hashing requires at most 2 MB of bytes.");
  return sha256HexDigest(bytes);
}

export {
  assertWorkbookStructureEditable,
  createUniqueWorksheetName,
  $v,
  validateDrawingColor,
  describeSheetDeletionImpacts,
  definedNameScopeKey,
  visitTableFormulasOutsideSheet,
  hashPictureBytes,
  maxPictureHashBytes,
  validateChartSettings,
  Kne,
  validateDrawingNumber,
  isNullOrEmptyValue,
  throwXlsxDefinedNameError,
  hashWorkbookStructurePassword,
  tableColumnDataTypes,
  spreadsheetMlNamespace,
  stripFormulaStringLiterals,
  coerceImportedValue,
  convertTextForDataType,
  Yne,
  validateXmlSafeName,
  formatDiagnosticsMessage,
  evaluateDefinedNameFunction,
  parseDefinedNameXmlAttributes,
  validateNameSheetReferences,
  rewriteStructuredTableReference,
  validateDefinedName,
  insertDefinedNamesIntoWorkbookXml,
  createUuid,
  replaceTextSpans,
  localBindingFunctionNames,
  parseDefinedNamesXml,
  validateTableColumnType,
  summarizeWorksheets,
  validateWorkbookPassword,
  m5,
  nP,
  validateWorkbookStructureProtection,
  supportedChartKinds,
  validateNameFormulaText,
  pO,
  evaluateReusableFormula,
  applyTableSourceEvidence,
  findWorksheetById,
  assertUniqueWorksheetName,
  formatQuotedSheetPrefix,
  reservedDefinedNames,
  parseWorksheetQualifiedReference,
  validateWorksheetVisibility,
  y5,
  validateTabColor,
  validateDrawingText,
  isVisibleWorksheet,
};
export { unquoteSheetName, quoteSheetName };

/* Also used by the workers. */
export {
  Bne,
  S5,
  inspectFormulaForMigration,
  isFormulaClosure,
  reusableFormulaLimits,
  sheetPrefixPattern,
};

/* Also used by the workers. */
export { invalidXmlCharPattern, isPlainObject, validateWorksheetName };
