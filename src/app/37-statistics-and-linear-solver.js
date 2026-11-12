import { buildCompositeKey } from "./11-local-workbook-store-and-versions.js";

import { fingerprintCanonicalValue } from "./01-schemas-date-functions-and-validation.js";
import {
  assertWorkbookStructureEditable,
  parseAnalysisRequest,
  FormulaParser,
  createWorksheetRecord,
  sumDecimalPartsToNumber,
  exactAverageOfNumbers,
  _t,
  parseRangeAddress,
  analysisLimits,
  evaluateRegressionFunction,
  analysisFormatVersion,
  validateDataValidationRule,
  resolveDefinedNameValue,
  normalizeWorkbookInput,
  columnIndexToLetters,
} from "./03-workbook-model.js";

import {
  readSourceHeaders,
  inspectOutputNameSafety,
} from "./12-consolidation-and-dependency-inspection.js";
import {
  assertSolverMagnitude,
  parseRangeBounds,
  assertAutomaticCalculation,
  normalizeAnalysisPermissions,
  rejectRangeAsScalar,
  p_,
  validateAnalysisSource,
} from "./36-watch-window-and-goal-seek.js";

function linearizeSolverFormula(workbook, analysisRequest, sourceOptions) {
  assertAutomaticCalculation(workbook);
  const variableCount = analysisRequest.variables.length,
    resolvedCells = new Map(),
    cellsInProgress = new Set();
  let syntaxNodeCount = 0;
  const constantForm = (constantValue) => ({
      constant: assertSolverMagnitude(constantValue),
      coefficients: Array(variableCount).fill(0),
    }),
    scaleForm = (linearForm, scalar) => ({
      constant: assertSolverMagnitude(linearForm.constant * scalar),
      coefficients: linearForm.coefficients.map((coefficient) =>
        assertSolverMagnitude(coefficient * scalar),
      ),
    }),
    addForms = (leftForm, rightForm) => ({
      constant: assertSolverMagnitude(leftForm.constant + rightForm.constant),
      coefficients: leftForm.coefficients.map((coefficient, coefficientIndex) =>
        assertSolverMagnitude(
          coefficient + rightForm.coefficients[coefficientIndex],
        ),
      ),
    }),
    multiplyForms = (firstForm, secondForm) => {
      if (
        firstForm.coefficients.some(Boolean) &&
        secondForm.coefficients.some(Boolean)
      )
        throw Error(
          "Nonlinear multiplication is outside the linear Solver subset.",
        );
      return firstForm.coefficients.some(Boolean)
        ? scaleForm(firstForm, secondForm.constant)
        : scaleForm(secondForm, firstForm.constant);
    },
    resolveCellForm = (sheetId, address, dependencyDepth = 0) => {
      const compositeKey = buildCompositeKey(sheetId, address);
      if (cellsInProgress.has(compositeKey))
        throw Error("Solver formulas contain a cycle.");
      if (resolvedCells.has(compositeKey))
        return resolvedCells.get(compositeKey);
      if (
        resolvedCells.size + cellsInProgress.size >=
          analysisLimits.dependencyCells ||
        dependencyDepth > analysisLimits.depth
      )
        throw Error("Solver dependency budget exceeded.");
      const sourceSheet = validateAnalysisSource(
          workbook,
          { sheetId: sheetId, address: address },
          sourceOptions,
        ),
        cell = sourceSheet.cells[address],
        variableIndex = analysisRequest.variables.findIndex(
          (variable) =>
            variable.cell.sheetId === sheetId &&
            variable.cell.address === address,
        );
      if (variableIndex >= 0) {
        const variableForm = constantForm(0);
        return (
          (variableForm.coefficients[variableIndex] = 1),
          resolvedCells.set(compositeKey, variableForm),
          variableForm
        );
      }
      if (!cell || cell.raw === null)
        throw Error(
          "Solver dependencies must be populated numeric cells; blanks and spill children are excluded.",
        );
      let resolvedForm;
      if (typeof cell.raw == "number") resolvedForm = constantForm(cell.raw);
      else if (
        typeof cell.raw == "string" &&
        cell.raw.startsWith("=") &&
        !cell.literal
      ) {
        if (cell.raw.length > analysisLimits.expressionCharacters)
          throw Error("Solver expression budget exceeded.");
        const formulaText = resolveDefinedNameValue(
          cell.raw,
          workbook,
          sheetId,
        ).slice(1);
        if (formulaText.length > 4096)
          throw Error("Expanded Solver expression budget exceeded.");
        (cellsInProgress.add(compositeKey),
          (resolvedForm = rejectRangeAsScalar(
            linearizeFormula(formulaText, sheetId, dependencyDepth + 1),
          )),
          cellsInProgress.delete(compositeKey));
      } else
        throw Error(
          "Solver dependencies must contain numbers or supported linear formulas.",
        );
      return (resolvedCells.set(compositeKey, resolvedForm), resolvedForm);
    },
    linearizeFormula = (expressionText, defaultSheetId, depth) => {
      const tokenPattern =
          /\s*(?:(\d*\.\d+(?:[Ee][+-]?\d+)?|\d+(?:\.\d*)?(?:[Ee][+-]?\d+)?)|((?:(?:'(?:[^']|'')+'|[A-Za-z_][A-Za-z0-9_.]*)!)?\$?[A-Za-z]{1,3}\$?[1-9]\d{0,6}(?::\$?[A-Za-z]{1,3}\$?[1-9]\d{0,6})?)(?![A-Za-z0-9_])|([A-Za-z_][A-Za-z0-9_.]*)|([()+\-*/,:%]))/gy,
        tokens = [];
      let position = 0;
      for (
        ;
        position < expressionText.length &&
        expressionText.slice(position).trim();
      ) {
        tokenPattern.lastIndex = position;
        const tokenMatch = tokenPattern.exec(expressionText);
        if (!tokenMatch || tokenMatch.index !== position)
          throw Error(
            "Unsupported Solver syntax. Use numeric references, arithmetic, SUM, AVERAGE or SUMPRODUCT.",
          );
        if (
          (tokens.push({
            type: tokenMatch[1] ? 1 : tokenMatch[2] ? 2 : tokenMatch[3] ? 3 : 4,
            text:
              tokenMatch[1] ?? tokenMatch[2] ?? tokenMatch[3] ?? tokenMatch[4],
          }),
          (position = tokenPattern.lastIndex),
          ++syntaxNodeCount > analysisLimits.syntaxNodes)
        )
          throw Error("Solver syntax budget exceeded.");
      }
      let tokenIndex = 0,
        nestingDepth = 0;
      const peekToken = () => tokens[tokenIndex]?.text,
        parseOperand = () => {
          if (++nestingDepth > 32)
            throw Error("Solver expression nesting limit exceeded.");
          let operandForm;
          const token = tokens[tokenIndex++];
          if (!token) throw Error("Incomplete Solver formula.");
          if (token.text === "+" || token.text === "-")
            operandForm = scaleForm(
              rejectRangeAsScalar(parseOperand()),
              token.text === "-" ? -1 : 1,
            );
          else if (token.type === 1)
            operandForm = constantForm(Number(token.text));
          else if (token.type === 2) {
            const separatorIndex = token.text.lastIndexOf("!"),
              sheetPrefix =
                separatorIndex < 0 ? null : token.text.slice(0, separatorIndex),
              cellAddressText = token.text
                .slice(separatorIndex + 1)
                .replace(/\$/g, "")
                .toUpperCase(),
              referencedSheetId =
                sheetPrefix === null
                  ? defaultSheetId
                  : workbook.sheets.find(
                      (sheet) =>
                        sheet.name.toLowerCase() ===
                        (sheetPrefix.startsWith("'")
                          ? sheetPrefix.slice(1, -1).replace(/''/g, "'")
                          : sheetPrefix
                        ).toLowerCase(),
                    )?.id;
            if (!referencedSheetId)
              throw Error("A Solver worksheet reference is missing.");
            if (cellAddressText.includes(":")) {
              const rangeAddress = parseRangeAddress(cellAddressText);
              if (
                (rangeAddress.lastRow - rangeAddress.firstRow + 1) *
                  (rangeAddress.lastCol - rangeAddress.firstCol + 1) >
                500
              )
                throw Error("Solver range exceeds 500 cells.");
              const rangeForms = [];
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
                  rangeForms.push(
                    resolveCellForm(
                      referencedSheetId,
                      `${columnIndexToLetters(firstCol)}${firstRow + 1}`,
                      depth,
                    ),
                  );
              operandForm = Object.assign(rangeForms, {
                height: rangeAddress.lastRow - rangeAddress.firstRow + 1,
                width: rangeAddress.lastCol - rangeAddress.firstCol + 1,
              });
            } else
              operandForm = resolveCellForm(
                referencedSheetId,
                cellAddressText,
                depth,
              );
          } else if (token.text === "(") {
            if (
              ((operandForm = parseSum()), tokens[tokenIndex++]?.text !== ")")
            )
              throw Error("Unbalanced Solver parentheses.");
          } else if (token.type === 3) {
            const functionName = token.text.toUpperCase();
            if (
              !["SUM", "AVERAGE", "SUMPRODUCT"].includes(functionName) ||
              tokens[tokenIndex++]?.text !== "("
            )
              throw Error(
                `Unsupported Solver function or name: ${token.text}.`,
              );
            const argumentForms = [];
            if (peekToken() !== ")")
              for (argumentForms.push(parseSum()); peekToken() === ",";)
                (tokenIndex++, argumentForms.push(parseSum()));
            if (tokens[tokenIndex++]?.text !== ")" || !argumentForms.length)
              throw Error("Invalid Solver function arguments.");
            if (functionName === "SUMPRODUCT") {
              const argumentRanges = argumentForms.map((argument) =>
                Array.isArray(argument) ? argument : [argument],
              );
              if (
                argumentRanges.some(
                  (argumentRange) =>
                    argumentRange.length !== argumentRanges[0].length ||
                    (argumentRange.height ?? 1) !==
                      (argumentRanges[0].height ?? 1) ||
                    (argumentRange.width ?? 1) !==
                      (argumentRanges[0].width ?? 1),
                )
              )
                throw Error("SUMPRODUCT ranges must have matching dimensions.");
              operandForm = argumentRanges[0].reduce(
                (accumulatedForm, rangeElement, cellIndex) =>
                  addForms(
                    accumulatedForm,
                    argumentRanges.reduce(
                      (rowProduct, factorRange) =>
                        multiplyForms(rowProduct, factorRange[cellIndex]),
                      constantForm(1),
                    ),
                  ),
                constantForm(0),
              );
            } else {
              const flatForms = argumentForms.flatMap((flatArgument) =>
                Array.isArray(flatArgument) ? flatArgument : [flatArgument],
              );
              ((operandForm = flatForms.reduce(addForms, constantForm(0))),
                functionName === "AVERAGE" &&
                  (operandForm = scaleForm(operandForm, 1 / flatForms.length)));
            }
          } else throw Error("Unsupported Solver expression.");
          for (; peekToken() === "%";)
            (tokenIndex++,
              (operandForm = scaleForm(
                rejectRangeAsScalar(operandForm),
                0.01,
              )));
          return (nestingDepth--, operandForm);
        },
        parseProduct = () => {
          let productResult = parseOperand();
          for (; peekToken() === "*" || peekToken() === "/";) {
            const text = tokens[tokenIndex++].text,
              rightOperandForm = rejectRangeAsScalar(parseOperand());
            if (text === "*")
              productResult = multiplyForms(
                rejectRangeAsScalar(productResult),
                rightOperandForm,
              );
            else {
              if (
                rightOperandForm.coefficients.some(Boolean) ||
                rightOperandForm.constant === 0
              )
                throw Error(
                  "Solver division requires a nonzero constant denominator.",
                );
              productResult = scaleForm(
                rejectRangeAsScalar(productResult),
                1 / rightOperandForm.constant,
              );
            }
          }
          return productResult;
        },
        parseSum = () => {
          let sumResult = parseProduct();
          for (; peekToken() === "+" || peekToken() === "-";) {
            const text = tokens[tokenIndex++].text,
              rightTermForm = rejectRangeAsScalar(parseProduct());
            sumResult = addForms(
              rejectRangeAsScalar(sumResult),
              scaleForm(rightTermForm, text === "-" ? -1 : 1),
            );
          }
          return sumResult;
        },
        rootForm = parseSum();
      if (tokenIndex !== tokens.length)
        throw Error("Unsupported or trailing Solver formula syntax.");
      return rootForm;
    },
    baseline = analysisRequest.variables.map((variable) => {
      const variableSheet = validateAnalysisSource(
          workbook,
          variable.cell,
          sourceOptions,
        ),
        cell = variableSheet.cells[variable.cell.address],
        variablePosition = _t(variable.cell.address);
      if (!cell || typeof cell.raw != "number" || !Number.isFinite(cell.raw))
        throw Error(
          "Changing cells must contain finite numeric inputs, not formulas or text.",
        );
      if (variableSheet.protection?.enabled)
        throw Error("A changing worksheet is protected.");
      if (
        [
          ...(variableSheet.pivots ?? []),
          ...(variableSheet.queries ?? []),
        ].some((generatedResult) =>
          parseRangeBounds(generatedResult.outputRange, variable.cell.address),
        )
      )
        throw Error("Generated results cannot be changing cells.");
      for (const table of variableSheet.tables ?? []) {
        if (!parseRangeBounds(table.range, variable.cell.address)) continue;
        const rangeAddress = parseRangeAddress(table.range);
        if (
          variablePosition.row === rangeAddress.firstRow ||
          (table.showTotals && variablePosition.row === rangeAddress.lastRow) ||
          table.columns[variablePosition.col - rangeAddress.firstCol]
            ?.calculatedFormula ||
          Object.entries(variableSheet.cells).some(([cellAddress, cell]) => {
            const otherCellPosition = _t(cellAddress);
            return (
              otherCellPosition.col === variablePosition.col &&
              otherCellPosition.row > rangeAddress.firstRow &&
              otherCellPosition.row <=
                rangeAddress.lastRow - +!!table.showTotals &&
              !cell.literal &&
              typeof cell.raw == "string" &&
              cell.raw.startsWith("=")
            );
          })
        )
          throw Error(
            "Table headers, totals and calculated columns cannot be changing cells.",
          );
      }
      if (
        (variableSheet.validations ?? []).some((validationRule) =>
          parseRangeBounds(validationRule.range, variable.cell.address),
        )
      )
        throw Error(
          "Changing cells with validation rules are excluded. Use explicit Solver bounds and constraints on a reviewed input copy.",
        );
      return cell.raw;
    }),
    objectiveSheet = validateAnalysisSource(
      workbook,
      analysisRequest.objective,
      sourceOptions,
    ),
    cell = objectiveSheet.cells[analysisRequest.objective.address];
  if (
    !cell ||
    typeof cell.raw != "string" ||
    !cell.raw.startsWith("=") ||
    cell.literal
  )
    throw Error("The objective must be a linear formula cell.");
  const objective = resolveCellForm(
      analysisRequest.objective.sheetId,
      analysisRequest.objective.address,
    ),
    constraints = analysisRequest.constraints.map((constraint) => ({
      form: addForms(
        resolveCellForm(constraint.left.sheetId, constraint.left.address),
        scaleForm(
          typeof constraint.right == "number"
            ? constantForm(constraint.right)
            : resolveCellForm(
                constraint.right.sheetId,
                constraint.right.address,
              ),
          -1,
        ),
      ),
      operator: constraint.operator,
    }));
  return (
    analysisRequest.direction === "target" &&
      constraints.push({
        form: addForms(objective, constantForm(-analysisRequest.target)),
        operator: "=",
      }),
    {
      objective: objective,
      constraints: constraints,
      baseline: baseline,
      dependencies: resolvedCells.size,
    }
  );
}
const evaluateLinearExpression = (linearForm, values) =>
  assertSolverMagnitude(
    linearForm.constant +
      linearForm.coefficients.reduce(
        (total, coefficient, coefficientIndex) =>
          total + coefficient * values[coefficientIndex],
        0,
      ),
  );
function solveLinearSystem(coefficientMatrix, rightHandSide) {
  const dimension = rightHandSide.length,
    augmented = coefficientMatrix.map((row, rowIndex) => {
      const rowScale = Math.max(...row.map(Math.abs));
      return [
        ...row.map((coefficient) => coefficient / (rowScale || 1)),
        rightHandSide[rowIndex] / (rowScale || 1),
      ];
    });
  for (let pivotColumn = 0; pivotColumn < dimension; pivotColumn++) {
    let pivotRow = pivotColumn;
    for (
      let candidateRow = pivotColumn + 1;
      candidateRow < dimension;
      candidateRow++
    )
      Math.abs(augmented[candidateRow][pivotColumn]) >
        Math.abs(augmented[pivotRow][pivotColumn]) && (pivotRow = candidateRow);
    const pivotMagnitude = Math.abs(augmented[pivotRow][pivotColumn]);
    if (pivotMagnitude < 1e-14) return "singular";
    if (pivotMagnitude < 1e-9) return "uncertain";
    [augmented[pivotRow], augmented[pivotColumn]] = [
      augmented[pivotColumn],
      augmented[pivotRow],
    ];
    const pivotValue = augmented[pivotColumn][pivotColumn];
    for (
      let normalizeColumn = pivotColumn;
      normalizeColumn <= dimension;
      normalizeColumn++
    )
      augmented[pivotColumn][normalizeColumn] /= pivotValue;
    for (let eliminationRow = 0; eliminationRow < dimension; eliminationRow++)
      if (eliminationRow !== pivotColumn) {
        const eliminationFactor = augmented[eliminationRow][pivotColumn];
        for (
          let eliminationColumn = pivotColumn;
          eliminationColumn <= dimension;
          eliminationColumn++
        )
          augmented[eliminationRow][eliminationColumn] -=
            eliminationFactor * augmented[pivotColumn][eliminationColumn];
      }
  }
  return augmented.map((augmentedRow) => augmentedRow[dimension]);
}
function solveBoundedLinearModel(workbook, analysisRequest, options = {}) {
  const linearModel = linearizeSolverFormula(
      workbook,
      analysisRequest,
      options,
    ),
    variableCount = analysisRequest.variables.length,
    isIntegerModel = analysisRequest.variables.every(
      (variable) => variable.domain !== "continuous",
    ),
    resultTemplate = {
      status: "infeasible",
      success: false,
      message:
        "No feasible solution satisfies the declared bounds and constraints.",
      algorithm: isIntegerModel
        ? "Exhaustive bounded integer enumeration"
        : "Exhaustive bounded linear vertex enumeration with pivoted elimination",
      assumptions: [
        "Symbolically proven linear formulas; explicit finite bounds on every changing cell.",
        "No nonlinear, evolutionary, mixed integer, unrestricted or proprietary Solver engine is used.",
        `Feasibility uses relative tolerance ${analysisRequest.tolerance}; all candidates must be exhausted before an optimal label.`,
        `Dependency cells: ${linearModel.dependencies}. No sample exclusions; all declared constraints participate.`,
        "Independent references: analytic linear fixtures and Python Fraction vertex checks; see ANALYSIS-VERIFICATION.json.",
      ],
      exclusions: [],
      included: variableCount,
      inspected: variableCount,
      iterations: 0,
      rows: [],
      baseline: linearModel.baseline,
    },
    inequalities = [];
  for (const constraint of linearModel.constraints) {
    const constraintRow = {
      a: constraint.form.coefficients,
      b: -constraint.form.constant,
    };
    (constraint.operator !== ">=" && inequalities.push(constraintRow),
      constraint.operator !== "<=" &&
        inequalities.push({
          a: constraintRow.a.map(
            (inequalityCoefficient) => -inequalityCoefficient,
          ),
          b: -constraintRow.b,
        }));
  }
  analysisRequest.variables.forEach((variable, variableIndex) => {
    const coefficients = Array(variableCount).fill(0);
    ((coefficients[variableIndex] = 1),
      inequalities.push(
        { a: coefficients, b: variable.upper },
        {
          a: coefficients.map((unitCoefficient) => -unitCoefficient),
          b: -variable.lower,
        },
      ));
  });
  const isFeasible = (candidate) =>
    candidate.every(Number.isFinite) &&
    inequalities.every((inequality) => {
      const leftHandSide = inequality.a.reduce(
          (total, coefficient, coefficientIndex) =>
            total + coefficient * candidate[coefficientIndex],
          0,
        ),
        scale = Math.max(
          1,
          Math.abs(inequality.b),
          inequality.a.reduce(
            (totalMagnitude, termCoefficient, termIndex) =>
              totalMagnitude + Math.abs(termCoefficient * candidate[termIndex]),
            0,
          ),
        );
      return leftHandSide - inequality.b <= analysisRequest.tolerance * scale;
    });
  let bestCandidate,
    objective = 0,
    hasPrecisionIssue = false,
    hasWorkLimit = false;
  const considerCandidate = (candidatePoint) => {
      if (!isFeasible(candidatePoint)) return;
      const candidateObjective = evaluateLinearExpression(
        linearModel.objective,
        candidatePoint,
      );
      (!bestCandidate ||
        (analysisRequest.direction === "maximize" &&
          candidateObjective > objective) ||
        (analysisRequest.direction === "minimize" &&
          candidateObjective < objective)) &&
        ((bestCandidate = [...candidatePoint]),
        (objective = candidateObjective));
    },
    canInspectCandidate = () =>
      resultTemplate.iterations >= analysisRequest.maxCandidates
        ? ((hasWorkLimit = true), false)
        : (resultTemplate.iterations++, true);
  if (isIntegerModel)
    if (
      analysisRequest.variables
        .map((variable) =>
          Math.max(
            0,
            Math.floor(variable.upper) - Math.ceil(variable.lower) + 1,
          ),
        )
        .reduce(
          (candidateCount, domainSize) => candidateCount * domainSize,
          1,
        ) > analysisRequest.maxCandidates
    )
      hasWorkLimit = true;
    else {
      const enumerateIntegerAssignments = (assignment) => {
        if (assignment.length === variableCount) {
          canInspectCandidate() && considerCandidate(assignment);
          return;
        }
        const variableSpec = analysisRequest.variables[assignment.length];
        for (
          let integerValue = Math.ceil(variableSpec.lower);
          integerValue <= Math.floor(variableSpec.upper);
          integerValue++
        )
          enumerateIntegerAssignments([...assignment, integerValue]);
      };
      enumerateIntegerAssignments([]);
    }
  else {
    let combinationCount = 1;
    for (
      let combinationIndex = 1;
      combinationIndex <= variableCount;
      combinationIndex++
    )
      combinationCount =
        (combinationCount *
          (inequalities.length - variableCount + combinationIndex)) /
        combinationIndex;
    if (combinationCount > analysisRequest.maxCandidates) hasWorkLimit = true;
    else {
      const enumerateCombinations = (chosenIndices, startIndex) => {
        if (chosenIndices.length === variableCount) {
          if (!canInspectCandidate()) return;
          const selectedInequalities = chosenIndices.map(
              (constraintIndex) => inequalities[constraintIndex],
            ),
            vertexPoint = solveLinearSystem(
              selectedInequalities.map(
                (selectedInequality) => selectedInequality.a,
              ),
              selectedInequalities.map((selectedRow) => selectedRow.b),
            );
          vertexPoint === "uncertain"
            ? (hasPrecisionIssue = true)
            : vertexPoint !== "singular" && considerCandidate(vertexPoint);
          return;
        }
        for (
          let nextConstraintIndex = startIndex;
          nextConstraintIndex <=
          inequalities.length - (variableCount - chosenIndices.length);
          nextConstraintIndex++
        )
          enumerateCombinations(
            [...chosenIndices, nextConstraintIndex],
            nextConstraintIndex + 1,
          );
      };
      enumerateCombinations([], 0);
    }
  }
  if (hasWorkLimit)
    return {
      ...resultTemplate,
      status: "work-limit",
      message:
        "The candidate budget is insufficient to certify a result. Reduce the model size or narrow integer bounds. Baseline unchanged.",
    };
  if (hasPrecisionIssue)
    return {
      ...resultTemplate,
      status: "precision-limit",
      message:
        "Near-dependent constraints prevent a reliable optimum certificate. Rescale or simplify the constraints. Baseline unchanged.",
    };
  if (!bestCandidate) return resultTemplate;
  const solution = bestCandidate.map((coordinate) =>
    Object.is(coordinate, -0) ? 0 : coordinate,
  );
  return (
    solution.forEach((solutionValue, solutionIndex) => {
      const cell = analysisRequest.variables[solutionIndex].cell;
      if (
        workbook.sheets
          .find((sheet) => sheet.id === cell.sheetId)
          .validations?.some(
            (validation) =>
              parseRangeBounds(validation.range, cell.address) &&
              !validateDataValidationRule(validation, solutionValue),
          )
      )
        throw Error("The solution violates a changing-cell validation rule.");
    }),
    {
      ...resultTemplate,
      status: "optimal",
      success: true,
      message:
        analysisRequest.direction === "target"
          ? "A feasible solution reaches the target within tolerance."
          : "All bounded candidates were examined; the best feasible objective is reported within tolerance.",
      solution: solution,
      objective: objective,
      rows: [
        ["Changing cell", "Baseline", "Solution"],
        ...analysisRequest.variables.map((variable, changingCellIndex) => [
          `${workbook.sheets.find((sheet) => sheet.id === variable.cell.sheetId).name}!${variable.cell.address}`,
          linearModel.baseline[changingCellIndex],
          solution[changingCellIndex],
        ]),
        [
          "Objective",
          evaluateLinearExpression(linearModel.objective, linearModel.baseline),
          objective,
        ],
        ["Constraint", "Operator", "Residual (left − right)"],
        ...linearModel.constraints.map((constraint, constraintPosition) => [
          `Constraint ${constraintPosition + 1}`,
          constraint.operator,
          evaluateLinearExpression(constraint.form, solution),
        ]),
      ],
    }
  );
}
const assertFiniteStatistic = (value) => {
    if (!Number.isFinite(value))
      throw Error(
        "The statistical result exceeds finite numeric precision. Rescale the sample.",
      );
    return value;
  },
  computeSampleVariance = (sample) => {
    const mean = exactAverageOfNumbers(sample);
    return assertFiniteStatistic(
      sumDecimalPartsToNumber(
        sample.map((sampleValue) => (sampleValue - mean) ** 2),
      ) /
        (sample.length - 1),
    );
  },
  distributionPrimitiveHost = new FormulaParser({
    onCell: () => {
      throw Error("Distribution primitives cannot read workbook cells.");
    },
    onRange: () => {
      throw Error("Distribution primitives cannot read workbook ranges.");
    },
  });
function evaluateDistributionPrimitive(primitiveName, primitiveArguments) {
  const primitiveResult = distributionPrimitiveHost.parse(
    `${primitiveName}(${primitiveArguments.map((argument) => (typeof argument == "boolean" ? String(argument).toUpperCase() : String(argument))).join(",")})`,
    { sheet: "Primitives", row: 1, col: 1 },
  );
  if (typeof primitiveResult != "number" || !Number.isFinite(primitiveResult))
    throw Error(`The ${primitiveName} distribution primitive failed.`);
  return primitiveResult;
}
function fitLinearRegression(responses, predictors) {
  const predictorCount = predictors[0].length;
  if (responses.length <= predictorCount + 1)
    throw Error("Regression needs more observations than fitted coefficients.");
  if (computeSampleVariance(responses) === 0)
    throw Error("Regression requires a nonconstant response.");
  const linestResult = evaluateRegressionFunction(
      "LINEST",
      [responses.map((response) => [response]), predictors, true, true],
      {
        retrieve: (cellValue) => cellValue,
        issue: (issueMessage) => new Error(issueMessage),
        isIssue: (possibleIssue) => possibleIssue instanceof Error,
      },
    ),
    residualDegrees = linestResult[3][1];
  if (residualDegrees !== responses.length - predictorCount - 1)
    throw Error(
      "Regression predictors are rank deficient or numerically dependent.",
    );
  const coefficient = [...linestResult[0]].reverse(),
    standardErrors = [...linestResult[1]].reverse(),
    fitted = evaluateRegressionFunction(
      "TREND",
      [
        responses.map((responseValue) => [responseValue]),
        predictors,
        predictors,
        true,
      ],
      {
        retrieve: (trendCellValue) => trendCellValue,
        issue: (trendIssueMessage) => new Error(trendIssueMessage),
        isIssue: (trendPossibleIssue) => trendPossibleIssue instanceof Error,
      },
    ).map((fittedRow) => fittedRow[0]),
    residuals = responses.map((observedResponse, observationIndex) =>
      assertFiniteStatistic(observedResponse - fitted[observationIndex]),
    );
  return {
    coefficient: coefficient,
    standardErrors: standardErrors,
    df: residualDegrees,
    r2: linestResult[2][0],
    standardError: linestResult[2][1],
    f: linestResult[3][0],
    ssRegression: linestResult[4][0],
    ssError: linestResult[4][1],
    fitted: fitted,
    residuals: residuals,
  };
}
function computeDescriptiveStatistics(workbook, analysisRequest, options = {}) {
  const rangeAddress = parseRangeAddress(analysisRequest.range),
    columnCount = rangeAddress.lastCol - rangeAddress.firstCol + 1,
    rowSpan = rangeAddress.lastRow - rangeAddress.firstRow;
  if (
    !/^[A-Z]{1,3}[1-9]\d{0,6}:[A-Z]{1,3}[1-9]\d{0,6}$/.test(
      analysisRequest.range,
    ) ||
    rowSpan < 1 ||
    rowSpan > 500 ||
    columnCount > 6 ||
    (rowSpan + 1) * columnCount > 2e3
  )
    throw Error(
      "Select a headed rectangle with 1–500 data rows, 1–6 columns and at most 2,000 cells.",
    );
  (validateAnalysisSource(
    workbook,
    {
      sheetId: analysisRequest.sheetId,
      address: analysisRequest.range.split(":")[0],
    },
    options,
  ),
    assertAutomaticCalculation(workbook));
  const sourceHeaders = readSourceHeaders(
      workbook,
      [{ sheetId: analysisRequest.sheetId, range: analysisRequest.range }],
      options,
    ),
    sourceSnapshot = sourceHeaders.snapshots[0];
  for (const dependencyCell of sourceHeaders.dependencyCells)
    validateAnalysisSource(
      workbook,
      { sheetId: dependencyCell.sheetId, address: dependencyCell.address },
      options,
    );
  const exclusions = [],
    completeRecords = [];
  for (const record of sourceSnapshot.records)
    if (
      record.values.some(
        (value) => typeof value != "number" || !Number.isFinite(value),
      )
    ) {
      if (analysisRequest.missing === "reject")
        throw Error(
          `Row ${record.row} includes blank or nonnumeric values. Correct it or explicitly exclude incomplete rows.`,
        );
      exclusions.push({
        row: record.row,
        locator: record.locator,
        reason:
          "Complete-case exclusion: at least one blank, text or boolean value.",
      });
    } else completeRecords.push({ ...record, values: record.values });
  const included = completeRecords.length,
    observations = completeRecords.map(
      (completeRecord) => completeRecord.values,
    ),
    headers = sourceSnapshot.headers,
    singleColumnValues = () => {
      if (columnCount !== 1)
        throw Error("This analysis requires exactly one numeric column.");
      return observations.map((observation) => observation[0]);
    };
  if (
    included <
    (analysisRequest.method === "histogram" ||
    analysisRequest.method === "distribution"
      ? 1
      : 2)
  )
    throw Error("Insufficient numeric observations after exclusions.");
  const assumptions = [
    "Headed numeric inputs; no text-to-number or date conversion. Complete-case exclusions apply to every selected column.",
    "Hidden cells require explicit authorization; calculation errors, dynamic/volatile formulas and generated report inputs are refused.",
    "Binary64 arithmetic; local MIT implementation and the existing pinned MIT fast-formula-parser distribution primitives. No proprietary engine or new license dependency.",
  ];
  let algorithm = "",
    rows = [];
  if (analysisRequest.method === "descriptive") {
    ((algorithm = "Two-pass sample moments and sorted inclusive quartiles"),
      assumptions.push(
        "Sample variance and standard deviation use n − 1. Median and quartiles use linear inclusive interpolation.",
        "Independent reference: Python statistics and analytic fixtures in ANALYSIS-VERIFICATION.json.",
      ));
    const statisticLabels = [
        "Count",
        "Sum",
        "Mean",
        "Standard error",
        "Median",
        "Sample standard deviation",
        "Sample variance",
        "Minimum",
        "Q1 (inclusive)",
        "Q3 (inclusive)",
        "Maximum",
        "Range",
      ],
      columnStatistics = headers.map((header, columnIndex) => {
        const columnValues = observations.map(
            (observationRow) => observationRow[columnIndex],
          ),
          sortedValues = [...columnValues].sort(
            (leftValue, rightValue) => leftValue - rightValue,
          ),
          sampleVariance = computeSampleVariance(columnValues),
          quantile = (fraction) => {
            const position = (included - 1) * fraction,
              lowerIndex = Math.floor(position);
            return (
              sortedValues[lowerIndex] +
              (sortedValues[Math.ceil(position)] - sortedValues[lowerIndex]) *
                (position - lowerIndex)
            );
          };
        return [
          included,
          assertFiniteStatistic(sumDecimalPartsToNumber(columnValues)),
          assertFiniteStatistic(exactAverageOfNumbers(columnValues)),
          assertFiniteStatistic(Math.sqrt(sampleVariance / included)),
          quantile(0.5),
          assertFiniteStatistic(Math.sqrt(sampleVariance)),
          sampleVariance,
          sortedValues[0],
          quantile(0.25),
          quantile(0.75),
          sortedValues[included - 1],
          assertFiniteStatistic(sortedValues[included - 1] - sortedValues[0]),
        ];
      });
    rows = [
      ["Statistic", ...headers],
      ...statisticLabels.map((statisticLabel, statisticIndex) => [
        statisticLabel,
        ...columnStatistics.map(
          (columnStatistic) => columnStatistic[statisticIndex],
        ),
      ]),
    ];
  } else if (
    analysisRequest.method === "correlation" ||
    analysisRequest.method === "covariance"
  ) {
    if (columnCount < 2)
      throw Error("Select at least two columns for a matrix.");
    ((algorithm =
      analysisRequest.method === "correlation"
        ? "Pearson product-moment correlation"
        : "Two-pass sample covariance"),
      assumptions.push(
        "Every matrix entry uses the same complete-case observations; covariance uses n − 1.",
        "Independent reference: analytic matrices and Python statistics in ANALYSIS-VERIFICATION.json.",
      ));
    const columnMeans = headers.map((header, meanColumnIndex) =>
        exactAverageOfNumbers(
          observations.map(
            (meanObservation) => meanObservation[meanColumnIndex],
          ),
        ),
      ),
      sumsOfSquares = headers.map((header, squareColumnIndex) =>
        assertFiniteStatistic(
          sumDecimalPartsToNumber(
            observations.map(
              (deviationObservation) =>
                (deviationObservation[squareColumnIndex] -
                  columnMeans[squareColumnIndex]) **
                2,
            ),
          ),
        ),
      );
    if (
      analysisRequest.method === "correlation" &&
      sumsOfSquares.some((sumOfSquares) => sumOfSquares === 0)
    )
      throw Error("Correlation is undefined for a constant column.");
    rows = [
      [analysisRequest.method, ...headers],
      ...headers.map((header, matrixRowIndex) => [
        header,
        ...headers.map((header, matrixColumnIndex) => {
          const crossProductSum = assertFiniteStatistic(
            sumDecimalPartsToNumber(
              observations.map(
                (covarianceObservation) =>
                  (covarianceObservation[matrixRowIndex] -
                    columnMeans[matrixRowIndex]) *
                  (covarianceObservation[matrixColumnIndex] -
                    columnMeans[matrixColumnIndex]),
              ),
            ),
          );
          return analysisRequest.method === "correlation"
            ? Math.max(
                -1,
                Math.min(
                  1,
                  assertFiniteStatistic(
                    crossProductSum /
                      Math.sqrt(sumsOfSquares[matrixRowIndex]) /
                      Math.sqrt(sumsOfSquares[matrixColumnIndex]),
                  ),
                ),
              )
            : assertFiniteStatistic(crossProductSum / (included - 1));
        }),
      ]),
    ];
  } else if (analysisRequest.method === "regression") {
    if (columnCount < 2)
      throw Error(
        "Regression requires Y in the first column and one or more X columns.",
      );
    const regressionFit = fitLinearRegression(
        observations.map((observationForY) => observationForY[0]),
        observations.map((observationForX) => observationForX.slice(1)),
      ),
      criticalT = evaluateDistributionPrimitive("T.INV.2T", [
        0.05,
        regressionFit.df,
      ]);
    ((algorithm =
      "Ordinary least squares with intercept; centered, scaled pivoted QR"),
      assumptions.push(
        "First column is Y; remaining columns are X. Independent errors, linear mean and constant error variance; t/F inference additionally assumes normal errors.",
        "No automatic causal interpretation. Rank deficient and insufficient samples are refused. Perfect fits have undefined t/F values shown explicitly.",
        "Independent certified reference: NIST StRD Norris, 36 observations; see ANALYSIS-VERIFICATION.json.",
      ));
    const coefficientRows = ["Intercept", ...headers.slice(1)].map(
        (coefficientLabel, coefficientIndex) => {
          const coefficientStandardError =
              regressionFit.standardErrors[coefficientIndex],
            coefficientEstimate = regressionFit.coefficient[coefficientIndex],
            tStatistic =
              coefficientStandardError === 0
                ? null
                : assertFiniteStatistic(
                    coefficientEstimate / coefficientStandardError,
                  );
          return [
            coefficientLabel,
            coefficientEstimate,
            coefficientStandardError,
            tStatistic,
            tStatistic === null
              ? null
              : evaluateDistributionPrimitive("BETA.DIST", [
                  regressionFit.df /
                    (regressionFit.df + tStatistic * tStatistic),
                  regressionFit.df / 2,
                  0.5,
                  true,
                ]),
            coefficientEstimate - criticalT * coefficientStandardError,
            coefficientEstimate + criticalT * coefficientStandardError,
          ];
        },
      ),
      fStatistic = typeof regressionFit.f == "number" ? regressionFit.f : null;
    rows = [
      ["Metric", "Value"],
      ["Observations", included],
      ["R squared", regressionFit.r2],
      ["Residual standard error", regressionFit.standardError],
      ["Residual df", regressionFit.df],
      ["Regression SS", regressionFit.ssRegression],
      ["Residual SS", regressionFit.ssError],
      ["F statistic", fStatistic],
      [
        "Significance F",
        fStatistic === null
          ? null
          : evaluateDistributionPrimitive("BETA.DIST", [
              regressionFit.df /
                (regressionFit.df + (columnCount - 1) * fStatistic),
              regressionFit.df / 2,
              (columnCount - 1) / 2,
              true,
            ]),
      ],
      [
        "Coefficient",
        "Estimate",
        "Standard error",
        "t",
        "Two-sided p",
        "95% lower",
        "95% upper",
      ],
      ...coefficientRows,
      ["Source row", "Observed Y", "Fitted Y", "Residual"],
      ...completeRecords.map((fitRecord, recordIndex) => [
        fitRecord.row,
        fitRecord.values[0],
        regressionFit.fitted[recordIndex],
        regressionFit.residuals[recordIndex],
      ]),
    ];
  } else if (analysisRequest.method === "histogram") {
    const numericValues = singleColumnValues(),
      bins = analysisRequest.bins;
    if (
      !bins?.length ||
      bins.some((bin, binIndex) => binIndex > 0 && bin <= bins[binIndex - 1])
    )
      throw Error(
        "Histogram bin upper bounds must be finite and strictly increasing.",
      );
    ((algorithm = "Explicit right-closed histogram bins"),
      assumptions.push(
        "First bin includes all values at or below its upper bound; later bins are (previous, upper]; last bin is overflow.",
        "Independent reference: exact integer counting fixtures in ANALYSIS-VERIFICATION.json.",
      ));
    const binCounts = Array(bins.length + 1).fill(0);
    numericValues.forEach((histogramValue) => {
      const matchedBinIndex = bins.findIndex((bin) => histogramValue <= bin);
      binCounts[matchedBinIndex < 0 ? bins.length : matchedBinIndex]++;
    });
    let cumulativeCount = 0;
    rows = [
      [
        "Upper bound",
        "Frequency",
        "Relative frequency",
        "Cumulative frequency",
      ],
      ...binCounts.map(
        (binCount, bucketIndex) => (
          (cumulativeCount += binCount),
          [
            bucketIndex === bins.length ? "Overflow" : bins[bucketIndex],
            binCount,
            binCount / included,
            cumulativeCount / included,
          ]
        ),
      ),
    ];
  } else if (analysisRequest.method === "distribution") {
    const points = singleColumnValues();
    ((algorithm =
      "Pinned analytic probability density/mass and cumulative distribution"),
      assumptions.push(
        "Selected cells are evaluation points, not a fitted sample. Distribution parameters are supplied explicitly.",
        "Independent reference: Python NormalDist, math.comb and factorial probability sums in ANALYSIS-VERIFICATION.json.",
      ));
    const distributionName = analysisRequest.distribution ?? "normal";
    let distributionArgs, distributionPrimitive;
    if (distributionName === "normal") {
      if (
        analysisRequest.mean === undefined ||
        analysisRequest.standardDeviation === undefined
      )
        throw Error(
          "Normal distribution needs a finite mean and positive standard deviation.",
        );
      ((distributionPrimitive = "NORM.DIST"),
        (distributionArgs = (normalPoint, isNormalCumulative) => [
          normalPoint,
          analysisRequest.mean,
          analysisRequest.standardDeviation,
          isNormalCumulative,
        ]),
        assumptions.push(
          `Normal mean ${analysisRequest.mean}; standard deviation ${analysisRequest.standardDeviation}.`,
        ));
    } else if (distributionName === "binomial") {
      if (
        analysisRequest.trials === undefined ||
        analysisRequest.probability === undefined ||
        points.some(
          (binomialCandidate) =>
            !Number.isInteger(binomialCandidate) ||
            binomialCandidate < 0 ||
            binomialCandidate > analysisRequest.trials,
        )
      )
        throw Error(
          "Binomial points must be integers from zero through the explicit trial count; provide p in [0,1].",
        );
      ((distributionPrimitive = "BINOM.DIST"),
        (distributionArgs = (binomialPoint, isBinomialCumulative) => [
          binomialPoint,
          analysisRequest.trials,
          analysisRequest.probability,
          isBinomialCumulative,
        ]),
        assumptions.push(
          `Independent trials ${analysisRequest.trials}; success probability ${analysisRequest.probability}.`,
        ));
    } else {
      if (
        analysisRequest.lambda === undefined ||
        points.some(
          (poissonCandidate) =>
            !Number.isInteger(poissonCandidate) ||
            poissonCandidate < 0 ||
            poissonCandidate > 1e3,
        )
      )
        throw Error(
          "Poisson points must be integers from zero through 1,000; provide a positive rate.",
        );
      ((distributionPrimitive = "POISSON.DIST"),
        (distributionArgs = (poissonPoint, isPoissonCumulative) => [
          poissonPoint,
          analysisRequest.lambda,
          isPoissonCumulative,
        ]),
        assumptions.push(
          `Poisson rate ${analysisRequest.lambda}; independent arrivals and constant rate.`,
        ));
    }
    rows = [
      [
        "Point",
        distributionName === "normal" ? "Density" : "Probability mass",
        "Cumulative probability",
      ],
      ...points.map((distributionPoint) => [
        distributionPoint,
        evaluateDistributionPrimitive(
          distributionPrimitive,
          distributionArgs(distributionPoint, false),
        ),
        evaluateDistributionPrimitive(
          distributionPrimitive,
          distributionArgs(distributionPoint, true),
        ),
      ]),
    ];
  } else {
    if (columnCount !== 2)
      throw Error(
        "Forecasting requires two columns: increasing equally spaced numeric period, then value.",
      );
    const periods = observations.map((periodRow) => periodRow[0]),
      observedValues = observations.map((valueRow) => valueRow[1]),
      periodStep = periods[1] - periods[0];
    if (
      !(periodStep > 0) ||
      periods.some(
        (period, periodIndex) =>
          periodIndex > 0 &&
          Math.abs(period - periods[periodIndex - 1] - periodStep) >
            1e-9 * Math.max(1, Math.abs(periodStep)),
      )
    )
      throw Error(
        "Forecast periods must be strictly increasing and equally spaced without gaps.",
      );
    const horizon = analysisRequest.horizon ?? 6,
      forecastMethod = analysisRequest.forecast ?? "linear",
      forecastPeriods = Array.from({ length: horizon }, (j, stepIndex) =>
        assertFiniteStatistic(
          periods[included - 1] + periodStep * (stepIndex + 1),
        ),
      ),
      forecastValues = [],
      fittedEntries = [];
    if (forecastMethod === "linear") {
      const trendFit = fitLinearRegression(
        observedValues,
        periods.map((trendPeriodInput) => [trendPeriodInput]),
      );
      algorithm = "Linear trend forecast using centered QR least squares";
      const trendRows = evaluateRegressionFunction(
        "TREND",
        [
          observedValues.map((trendValueInput) => [trendValueInput]),
          periods.map((trendPeriodArg) => [trendPeriodArg]),
          forecastPeriods.map((trendForecastInput) => [trendForecastInput]),
          true,
        ],
        {
          retrieve: (retrievedValue) => retrievedValue,
          issue: (issueMessage) => new Error(issueMessage),
          isIssue: (rawResult) => rawResult instanceof Error,
        },
      );
      (forecastValues.push(...trendRows.map((trendRow) => trendRow[0])),
        trendFit.fitted.forEach((fittedValue, fittedIndex) =>
          fittedEntries.push({ index: fittedIndex, value: fittedValue }),
        ));
    } else if (forecastMethod === "moving-average") {
      const windowSize = analysisRequest.window ?? 3;
      if (windowSize >= included)
        throw Error("Moving average needs more observations than its window.");
      algorithm = `Trailing ${windowSize}-period moving average with recursive forecasts`;
      for (
        let observationIndex = windowSize;
        observationIndex < included;
        observationIndex++
      )
        fittedEntries.push({
          index: observationIndex,
          value: exactAverageOfNumbers(
            observedValues.slice(
              observationIndex - windowSize,
              observationIndex,
            ),
          ),
        });
      const extendedValues = [...observedValues];
      for (let horizonIndex = 0; horizonIndex < horizon; horizonIndex++) {
        const forecastValue = assertFiniteStatistic(
          exactAverageOfNumbers(extendedValues.slice(-windowSize)),
        );
        (forecastValues.push(forecastValue),
          extendedValues.push(forecastValue));
      }
    } else {
      if (included < 3)
        throw Error("Exponential smoothing needs at least three observations.");
      const smoothingAlpha = analysisRequest.alpha ?? 0.3;
      algorithm =
        "Simple exponential smoothing with fixed alpha and first-observation initialization";
      let smoothedValue = observedValues[0];
      for (let smoothingIndex = 1; smoothingIndex < included; smoothingIndex++)
        (fittedEntries.push({ index: smoothingIndex, value: smoothedValue }),
          (smoothedValue = assertFiniteStatistic(
            smoothingAlpha * observedValues[smoothingIndex] +
              (1 - smoothingAlpha) * smoothedValue,
          )));
      (forecastValues.push(...forecastPeriods.map(() => smoothedValue)),
        assumptions.push(
          `Alpha ${smoothingAlpha}; no trend or seasonal component.`,
        ));
    }
    const residuals = fittedEntries.map((fittedEntry) =>
      assertFiniteStatistic(
        observedValues[fittedEntry.index] - fittedEntry.value,
      ),
    );
    (assumptions.push(
      "Numeric equally spaced periods; no seasonality, automatic tuning, ETS intervals or causal claims. Extrapolation can be unreliable.",
      "Error metrics use fitted/in-sample observations, not an independent holdout. No uncertainty interval is claimed.",
      "Independent reference: analytic forecasts and Python calculations in ANALYSIS-VERIFICATION.json.",
    ),
      (rows = [
        ["Metric", "Value"],
        ["Fitted observations", fittedEntries.length],
        ["In-sample MAE", exactAverageOfNumbers(residuals.map(Math.abs))],
        [
          "In-sample RMSE",
          Math.sqrt(
            exactAverageOfNumbers(
              residuals.map((residualValue) =>
                assertFiniteStatistic(residualValue * residualValue),
              ),
            ),
          ),
        ],
        ["Period", "Forecast"],
        ...forecastPeriods.map((forecastPeriod, forecastIndex) => [
          forecastPeriod,
          assertFiniteStatistic(forecastValues[forecastIndex]),
        ]),
        ["Source row", "Period", "Observed", "Fitted", "Residual"],
        ...fittedEntries.map((fittedObservation) => [
          completeRecords[fittedObservation.index].row,
          periods[fittedObservation.index],
          observedValues[fittedObservation.index],
          fittedObservation.value,
          observedValues[fittedObservation.index] - fittedObservation.value,
        ]),
      ]));
  }
  for (const row of rows)
    for (const cell of row)
      typeof cell == "number" && assertFiniteStatistic(cell);
  return {
    status: "complete",
    success: true,
    message:
      "Calculation completed for the declared method and included observations. Review the assumptions and exclusions before saving.",
    algorithm: algorithm,
    assumptions: assumptions,
    exclusions: exclusions,
    included: included,
    inspected: sourceSnapshot.records.length,
    iterations: 0,
    rows: rows,
  };
}
const analysisResultCache = new WeakMap();
function runAnalysisRequest(workbook, request, options = {}) {
  const analysisRequest = parseAnalysisRequest(request),
    analysisPermissions = normalizeAnalysisPermissions(workbook, options),
    analysisResult =
      analysisRequest.kind === "solver"
        ? solveBoundedLinearModel(workbook, analysisRequest, options)
        : computeDescriptiveStatistics(workbook, analysisRequest, options),
    analysisRecord = {
      request: analysisRequest,
      result: analysisResult,
      workbookHash: fingerprintCanonicalValue(workbook),
      contextHash: fingerprintCanonicalValue(analysisPermissions),
      sourceHash: p_(workbook),
    };
  return (
    analysisResultCache.set(
      analysisRecord,
      fingerprintCanonicalValue(analysisRecord),
    ),
    analysisRecord
  );
}
function buildAnalysisReport(analysisRecord) {
  const result = analysisRecord.result;
  return [
    [
      analysisRecord.request.kind === "solver"
        ? "Solver report"
        : "Statistical analysis report",
    ],
    ["Outcome", result.status],
    ["Message", result.message],
    ["Algorithm", result.algorithm],
    ["Engine", analysisFormatVersion],
    [
      "License",
      "MIT; existing MIT distribution dependency; no proprietary engine",
    ],
    ["Inspected", result.inspected],
    ["Included", result.included],
    ["Excluded", result.exclusions.length],
    ["Candidates examined", result.iterations],
    ...result.assumptions.map((assumption, assumptionIndex) => [
      `Assumption ${assumptionIndex + 1}`,
      assumption,
    ]),
    [],
    ...result.rows,
    [],
    ["Excluded source row", "Reason", "Locator"],
    ...result.exclusions.map((exclusion) => [
      exclusion.row,
      exclusion.reason,
      exclusion.locator,
    ]),
  ];
}
function buildAnalysisReportCells(worksheetRecord, outputRange) {
  const rangeAddress = parseRangeAddress(outputRange);
  if (
    (rangeAddress.lastRow - rangeAddress.firstRow + 1) *
      (rangeAddress.lastCol - rangeAddress.firstCol + 1) >
    2e4
  )
    throw Error("Analysis output exceeds 20,000 cells.");
  const cellSnapshots = [];
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
      const cell =
        worksheetRecord.cells[
          `${columnIndexToLetters(firstCol)}${firstRow + 1}`
        ];
      cellSnapshots.push({
        raw: cell?.raw ?? null,
        literal: !!cell?.literal,
      });
    }
  return fingerprintCanonicalValue(cellSnapshots);
}
function commitAnalysisResult(
  workbook,
  analysisRecord,
  commitMode,
  options = {},
) {
  const analysisPermissions = normalizeAnalysisPermissions(workbook, options);
  if (analysisPermissions.readOnly) throw Error("This workbook is read-only.");
  if (
    (assertWorkbookStructureEditable(workbook),
    !["report", "apply-solution"].includes(commitMode))
  )
    throw Error("Choose Save report or Apply solution and report.");
  if (
    analysisResultCache.get(analysisRecord) !==
      fingerprintCanonicalValue(analysisRecord) ||
    analysisRecord.workbookHash !== fingerprintCanonicalValue(workbook) ||
    analysisRecord.contextHash !==
      fingerprintCanonicalValue(analysisPermissions)
  )
    throw Error(
      "The preview, workbook or host permissions changed. Preview again.",
    );
  if (!analysisRecord.result.success)
    throw Error("Only a successful reviewed analysis can be saved or applied.");
  if (
    commitMode === "apply-solution" &&
    (analysisRecord.request.kind !== "solver" ||
      !analysisRecord.result.solution)
  )
    throw Error("Only Solver can apply changing-cell values.");
  if (
    workbook.sheets.some(
      (sheet) =>
        sheet.name.toLowerCase() ===
        analysisRecord.request.outputName.toLowerCase(),
    )
  )
    throw Error("The output worksheet name is already in use.");
  inspectOutputNameSafety(workbook, analysisRecord.request.outputName);
  const workbookInput = normalizeWorkbookInput(workbook);
  commitMode === "apply-solution" &&
    analysisRecord.request.kind === "solver" &&
    analysisRecord.request.variables.forEach((variable, variableIndex) => {
      const sheet = workbookInput.sheets.find(
        (sheet) => sheet.id === variable.cell.sheetId,
      );
      sheet.cells[variable.cell.address] = {
        ...sheet.cells[variable.cell.address],
        raw: analysisRecord.result.solution[variableIndex],
      };
    });
  const worksheetRecord = createWorksheetRecord(
      analysisRecord.request.outputName,
    ),
    analysisReport = buildAnalysisReport(analysisRecord),
    columnCount = Math.max(...analysisReport.map((row) => row.length));
  if (analysisReport.length * columnCount > 2e4)
    throw Error("Analysis report exceeds the output budget.");
  (analysisReport.forEach((reportRow, rowIndex) =>
    reportRow.forEach((raw, columnIndex) => {
      worksheetRecord.cells[
        `${columnIndexToLetters(columnIndex)}${rowIndex + 1}`
      ] = {
        raw: raw,
        ...(typeof raw == "string" ? { literal: true } : {}),
        ...(rowIndex === 0
          ? { style: { bold: true, fill: "#362C63", color: "#FFFFFF" } }
          : {}),
      };
    }),
  ),
    (worksheetRecord.columnWidths = {
      0: 210,
      1: 260,
      2: 180,
      3: 160,
      4: 140,
      5: 140,
      6: 140,
    }),
    (worksheetRecord.frozenRows = 1));
  const outputRange = `A1:${columnIndexToLetters(columnCount - 1)}${analysisReport.length}`;
  ((worksheetRecord.analysisReport = {
    engine: analysisFormatVersion,
    request: structuredClone(analysisRecord.request),
    result: structuredClone(analysisRecord.result),
    definitionHash: fingerprintCanonicalValue(analysisRecord.request),
    sourceHash: p_(workbookInput),
    outputHash: buildAnalysisReportCells(worksheetRecord, outputRange),
    outputRange: outputRange,
    createdAt: new Date().toISOString(),
    sourceRevision: workbook.revision,
    includeHidden: !!options.includeHidden,
    applied: commitMode === "apply-solution",
  }),
    workbookInput.sheets.push(worksheetRecord));
  const committedWorkbook = normalizeWorkbookInput(workbookInput);
  return (analysisResultCache.delete(analysisRecord), committedWorkbook);
}
function readSavedAnalysisReport(workbook, sheetId, options = {}) {
  const sheet = workbook.sheets.find((sheet) => sheet.id === sheetId),
    savedReport = sheet?.analysisReport;
  if (!sheet || !savedReport)
    return { status: "unavailable", message: "No saved analysis report." };
  const analysisPermissions = normalizeAnalysisPermissions(workbook, options);
  if (
    !analysisPermissions.allowedSheetIds.includes(sheetId) ||
    (!analysisPermissions.includeHidden && (sheet.hidden || sheet.veryHidden))
  )
    return {
      status: "unavailable",
      message: "The report is outside the authorized context.",
    };
  if (
    (savedReport.includeHidden && !analysisPermissions.includeHidden) ||
    analysisPermissions.allowedSheetIds.length < workbook.sheets.length
  )
    return {
      status: "unavailable",
      message:
        "Reopen the original analysis within the authorized source context before trusting this historical report.",
    };
  if (
    savedReport.definitionHash !==
      fingerprintCanonicalValue(savedReport.request) ||
    savedReport.sourceHash !== p_(workbook) ||
    savedReport.outputHash !==
      buildAnalysisReportCells(sheet, savedReport.outputRange)
  )
    return {
      status: "stale",
      message:
        "Source data, calculation settings or report cells changed. Preview the analysis again; saved values are historical.",
    };
  const worksheetRecord = createWorksheetRecord("Expected");
  return (
    buildAnalysisReport(savedReport).forEach((reportRow, rowIndex) =>
      reportRow.forEach(
        (raw, columnIndex) =>
          (worksheetRecord.cells[
            `${columnIndexToLetters(columnIndex)}${rowIndex + 1}`
          ] = {
            raw: raw,
            ...(typeof raw == "string" ? { literal: true } : {}),
          }),
      ),
    ),
    buildAnalysisReportCells(worksheetRecord, savedReport.outputRange) !==
    savedReport.outputHash
      ? {
          status: "stale",
          message: "The saved result and report cells disagree. Preview again.",
        }
      : {
          status: "current",
          message:
            "Historical reviewed snapshot; source fingerprint is unchanged. Reopen to rerun or refresh.",
        }
  );
}
const parseNumericText = (text) => (text.trim() === "" ? NaN : Number(text));

export {
  readSavedAnalysisReport,
  commitAnalysisResult,
  runAnalysisRequest,
  parseNumericText,
};
