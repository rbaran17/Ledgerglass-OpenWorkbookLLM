import i from "react/jsx-runtime";
import ne from "react";
import { fingerprintCanonicalValue } from "./01-schemas-date-functions-and-validation.js";
import {
  assertWorkbookStructureEditable,
  daxFunctionNames,
  walkDaxSyntaxTree,
  createWorksheetRecord,
  DaxError,
  columnarModelFormatVersion,
  formatOccupiedRangeAddress,
  parseDaxExpression,
  validateModelQuery,
  validateDataModel,
  parseRangeAddress,
  throwDaxError,
  normalizeWorkbookInput,
  daxExpressionLimits,
  dataModelLimits,
  columnIndexToLetters,
} from "./03-workbook-model.js";
import {
  useWorkbookMessages,
  LocalizedText,
} from "./04-cell-input-paste-and-fill.js";
import {
  readSourceHeaders,
  inspectOutputNameSafety,
} from "./12-consolidation-and-dependency-inspection.js";
import {
  requireNumericOperand,
  requireLogicalOperand,
  buildTableColumnKey,
  requireScalarValue,
  requireFiniteNumber,
  sumWithCompensation,
  compareDaxValues,
  caseFoldText,
  buildTypedValueKey,
} from "./13-consolidation-dialog.js";

function evaluateDaxExpression(dataAccess) {
  const definition = dataAccess.definition,
    tablesById = new Map(definition.tables.map((table) => [table.id, table])),
    parsedExpressions = new Map(),
    calculatedValueCache = new Map(),
    activeRowEvaluations = new Set(),
    measureStack = [],
    filterMaskCache = new WeakMap(),
    findTable = (tableName) =>
      definition.tables.find(
        (table) => caseFoldText(table.name) === caseFoldText(tableName),
      ) ?? throwDaxError("DAX_REFERENCE", `Unknown model table ${tableName}.`),
    resolveColumn = (columnRef) => {
      if (columnRef.kind !== "field")
        return throwDaxError(
          "DAX_ARGUMENT",
          "A fully qualified model column is required.",
        );
      const ownerTable = findTable(columnRef.table),
        column = ownerTable.columns.find(
          (column) =>
            caseFoldText(column.name) === caseFoldText(columnRef.column),
        );
      return column
        ? { tableId: ownerTable.id, column: column.name }
        : throwDaxError(
            "DAX_REFERENCE",
            `Unknown model column ${columnRef.table}[${columnRef.column}].`,
          );
    },
    findMeasure = (measureName) =>
      definition.measures.find(
        (measure) => caseFoldText(measure.name) === caseFoldText(measureName),
      ) ??
      throwDaxError(
        "DAX_REFERENCE",
        `Unknown measure [${measureName}]. Use qualified table[column] references for columns.`,
      );
  for (const table of definition.tables)
    for (const column of table.columns)
      column.expression !== undefined &&
        parsedExpressions.set(
          buildTableColumnKey({ tableId: table.id, column: column.name }),
          parseDaxExpression(column.expression),
        );
  for (const measure of definition.measures)
    measure.operation === "dax" &&
      parsedExpressions.set(
        `measure:${measure.id}`,
        parseDaxExpression(measure.expression),
      );
  for (const parsedExpression of parsedExpressions.values())
    walkDaxSyntaxTree(parsedExpression, (walkNode) => {
      (walkNode.kind === "field" && resolveColumn(walkNode),
        walkNode.kind === "table" && findTable(walkNode.name),
        walkNode.kind === "measure" && findMeasure(walkNode.name));
    });
  const isTableArgument = (argumentNode) =>
      argumentNode.kind === "table" ||
      (argumentNode.kind === "call" &&
        ["FILTER", "ALL"].includes(argumentNode.name)),
    isNumericLiteralArgument = (candidateNode) =>
      (candidateNode.kind === "literal" &&
        typeof candidateNode.value == "number") ||
      (candidateNode.kind === "unary" &&
        ["+", "-"].includes(candidateNode.op) &&
        candidateNode.value.kind === "literal" &&
        typeof candidateNode.value.value == "number") ||
      (candidateNode.kind === "call" && candidateNode.name === "BLANK");
  for (const modelExpression of parsedExpressions.values())
    walkDaxSyntaxTree(modelExpression, (callNode) => {
      if (callNode.kind !== "call") return;
      const requireQualifiedArgument = (argumentIndex) => {
        callNode.args[argumentIndex].kind !== "field" &&
          throwDaxError(
            "DAX_ARGUMENT",
            `${callNode.name} requires a qualified table[column] reference.`,
          );
      };
      if (
        ([
          "SUM",
          "AVERAGE",
          "MIN",
          "MAX",
          "DISTINCTCOUNT",
          "SELECTEDVALUE",
          "RELATED",
        ].includes(callNode.name) && requireQualifiedArgument(0),
        ["SUM", "AVERAGE", "MIN", "MAX"].includes(callNode.name))
      ) {
        const aggregateColumn = resolveColumn(callNode.args[0]);
        tablesById
          .get(aggregateColumn.tableId)
          .columns.find((column) => column.name === aggregateColumn.column)
          .type !== "number" &&
          throwDaxError(
            "DAX_TYPE",
            `${callNode.name} requires a number column in this subset.`,
          );
      }
      (callNode.name === "TOTALYTD" && requireQualifiedArgument(1),
        callNode.name === "ALL" &&
          !["table", "field"].includes(callNode.args[0].kind) &&
          throwDaxError(
            "DAX_ARGUMENT",
            "ALL accepts one table or qualified column.",
          ),
        ["FILTER", "SUMX", "COUNTROWS"].includes(callNode.name) &&
          !isTableArgument(callNode.args[0]) &&
          throwDaxError(
            "DAX_ARGUMENT",
            `${callNode.name} requires a table, FILTER or ALL expression.`,
          ),
        callNode.name === "DIVIDE" &&
          callNode.args[2] &&
          !isNumericLiteralArgument(callNode.args[2]) &&
          throwDaxError(
            "DAX_ARGUMENT",
            "DIVIDE alternate result must be a numeric literal or BLANK().",
          ),
        callNode.args.forEach((arg, argumentPosition) => {
          if (!isTableArgument(arg)) return;
          (["FILTER", "SUMX", "COUNTROWS", "ALL"].includes(callNode.name) &&
            argumentPosition === 0) ||
            (callNode.name === "CALCULATE" && argumentPosition > 0) ||
            throwDaxError(
              "DAX_ARGUMENT",
              `${callNode.name} requires a scalar at argument ${argumentPosition + 1}.`,
            );
        }));
    });
  const collectDependencies = (sourceExpression) => {
      const dependencyKeys = [];
      return (
        walkDaxSyntaxTree(sourceExpression, (dependencyNode) => {
          (dependencyNode.kind === "field" &&
            dependencyKeys.push(
              buildTableColumnKey(resolveColumn(dependencyNode)),
            ),
            dependencyNode.kind === "measure" &&
              dependencyKeys.push(
                `measure:${findMeasure(dependencyNode.name).id}`,
              ));
        }),
        dependencyKeys
      );
    },
    resolvedDependencies = new Set(),
    inProgressDependencies = new Set(),
    visitDependency = (dependencyKey) => {
      if (
        (inProgressDependencies.has(dependencyKey) &&
          throwDaxError(
            "DAX_CYCLE",
            "Calculated columns or measures contain a circular dependency.",
          ),
        resolvedDependencies.has(dependencyKey))
      )
        return;
      const dependencyExpression = parsedExpressions.get(dependencyKey);
      dependencyExpression &&
        (inProgressDependencies.add(dependencyKey),
        collectDependencies(dependencyExpression).forEach(visitDependency),
        inProgressDependencies.delete(dependencyKey),
        resolvedDependencies.add(dependencyKey));
    };
  parsedExpressions.forEach((N, expressionKey) =>
    visitDependency(expressionKey),
  );
  const allRowIndexes = (sourceTableId) =>
      Array.from(
        { length: dataAccess.count(sourceTableId) },
        (R, rowIndex) => rowIndex,
      ),
    readColumnValue = (columnReference, targetRowIndex) => {
      dataAccess.tick();
      const columnTable = tablesById.get(columnReference.tableId),
        column = columnTable.columns.find(
          (column) => column.name === columnReference.column,
        );
      if (column.expression === undefined)
        return dataAccess.value(
          columnReference.tableId,
          columnReference.column,
          targetRowIndex,
        );
      const tableColumnKey = buildTableColumnKey(columnReference),
        rowCacheKey = `${tableColumnKey}:${targetRowIndex}`;
      let cachedRowValues = calculatedValueCache.get(tableColumnKey);
      if (
        (cachedRowValues ||
          ((cachedRowValues = Array(dataAccess.count(columnReference.tableId))),
          calculatedValueCache.set(tableColumnKey, cachedRowValues)),
        Object.prototype.hasOwnProperty.call(cachedRowValues, targetRowIndex))
      )
        return cachedRowValues[targetRowIndex];
      if (
        activeRowEvaluations.has(rowCacheKey) ||
        activeRowEvaluations.size > 48
      )
        return throwDaxError(
          "DAX_CYCLE",
          "A calculated column has a circular or excessively deep evaluation dependency.",
        );
      activeRowEvaluations.add(rowCacheKey);
      try {
        const computedValue = requireScalarValue(
          evaluateNode(parsedExpressions.get(tableColumnKey), {
            filters: [],
            rows: [{ tableId: columnReference.tableId, index: targetRowIndex }],
          }),
        );
        return (
          isValidColumnValue(computedValue, column) ||
            throwDaxError(
              "DAX_TYPE",
              `${columnTable.name}[${column.name}] row ${targetRowIndex + 1} does not match its declared ${column.type}${column.nullable ? " or BLANK" : ""} type.`,
            ),
          (cachedRowValues[targetRowIndex] = computedValue),
          computedValue
        );
      } finally {
        activeRowEvaluations.delete(rowCacheKey);
      }
    },
    buildTableMasks = (filterState) => {
      const cachedMasks = filterMaskCache.get(filterState.filters);
      if (cachedMasks) return cachedMasks;
      const tableMasks = new Map(
        definition.tables.map((table) => [
          table.id,
          new Uint8Array(dataAccess.count(table.id)).fill(1),
        ]),
      );
      for (const filter of filterState.filters) {
        const filterMask = tableMasks.get(filter.tableId);
        for (
          let maskRowIndex = 0;
          maskRowIndex < filterMask.length;
          maskRowIndex++
        )
          (dataAccess.tick(),
            filterMask[maskRowIndex] &&
              !filter.test(maskRowIndex) &&
              (filterMask[maskRowIndex] = 0));
      }
      for (const orderedTableId of dataAccess.order)
        for (const oneSideRelationship of definition.relationships.filter(
          (relationship) => relationship.one.tableId === orderedTableId,
        )) {
          const oneSideMask = tableMasks.get(orderedTableId),
            manySideMask = tableMasks.get(oneSideRelationship.many.tableId);
          for (
            let manyRowIndex = 0;
            manyRowIndex < manySideMask.length;
            manyRowIndex++
          ) {
            dataAccess.tick();
            const oneSideRowIndex = dataAccess.parentRow(
              oneSideRelationship.many.tableId,
              orderedTableId,
              manyRowIndex,
            );
            (oneSideRowIndex === undefined || !oneSideMask[oneSideRowIndex]) &&
              (manySideMask[manyRowIndex] = 0);
          }
        }
      return (filterMaskCache.set(filterState.filters, tableMasks), tableMasks);
    },
    visibleRows = (visibleTableId, contextState) =>
      allRowIndexes(visibleTableId).filter(
        (candidateRowIndex) =>
          buildTableMasks(contextState).get(visibleTableId)[candidateRowIndex],
      ),
    stripFilterColumns = (filterList, excludedColumnKeys) =>
      filterList.flatMap((filterItem) => {
        const remainingColumns = filterItem.columns.filter(
          (column) =>
            !excludedColumnKeys.has(
              buildTableColumnKey({
                tableId: filterItem.tableId,
                column: column,
              }),
            ),
        );
        return remainingColumns.length === filterItem.columns.length
          ? [filterItem]
          : remainingColumns.length
            ? filterItem.project
              ? [filterItem.project(remainingColumns)]
              : throwDaxError(
                  "DAX_FILTER",
                  "This compound filter cannot be partially replaced. Use single-column filters.",
                )
            : [];
      }),
    mergeFilterSets = (existingFilters, newFilters) => {
      const newColumnKeys = new Set(
        newFilters.flatMap((newFilter) =>
          newFilter.columns.map((column) =>
            buildTableColumnKey({ tableId: newFilter.tableId, column: column }),
          ),
        ),
      );
      return [
        ...stripFilterColumns(existingFilters, newColumnKeys),
        ...newFilters,
      ];
    },
    tablesAffectedBy = (filteringTableId) =>
      definition.tables.filter((table) =>
        dataAccess.canFilter(table.id, filteringTableId),
      ),
    createRowSetFilter = (tableId, columns, matchedRowIndexes) => {
      const matchedRowKeys = matchedRowIndexes.map((matchedRowIndex) =>
          columns.map((column) =>
            buildTypedValueKey(
              readColumnValue(
                { tableId: tableId, column: column },
                matchedRowIndex,
              ),
            ),
          ),
        ),
        matchedKeySet = new Set(
          matchedRowKeys.map((rowKeyParts) => JSON.stringify(rowKeyParts)),
        );
      return {
        tableId: tableId,
        columns: columns,
        test: (testedRowIndex) =>
          matchedKeySet.has(
            JSON.stringify(
              columns.map((column) =>
                buildTypedValueKey(
                  readColumnValue(
                    { tableId: tableId, column: column },
                    testedRowIndex,
                  ),
                ),
              ),
            ),
          ),
        project: (projectedColumns) =>
          createRowSetFilter(tableId, projectedColumns, matchedRowIndexes),
      };
    },
    buildRowFilters = (rowFilterRequest) =>
      (rowFilterRequest.columns
        ? [tablesById.get(rowFilterRequest.tableId)]
        : tablesAffectedBy(rowFilterRequest.tableId)
      ).map((targetTable) => {
        const filterColumnNames =
            rowFilterRequest.columns ??
            targetTable.columns
              .filter((column) => column.expression === undefined)
              .map((plainColumn) => plainColumn.name),
          relatedRowIndexes = rowFilterRequest.rows
            .map((row) =>
              dataAccess.parentRow(
                rowFilterRequest.tableId,
                targetTable.id,
                row,
              ),
            )
            .filter((definedRowIndex) => definedRowIndex !== undefined);
        return createRowSetFilter(
          targetTable.id,
          filterColumnNames,
          relatedRowIndexes,
        );
      }),
    transitionRowContext = (rowContextState) => {
      if (rowContextState.transitioned || !rowContextState.rows.length)
        return rowContextState;
      let filters = [...rowContextState.filters];
      for (const row of rowContextState.rows)
        filters = mergeFilterSets(
          filters,
          buildRowFilters({
            tableId: row.tableId,
            rows: [row.index],
            columns: row.columns,
          }),
        );
      return { ...rowContextState, filters: filters, transitioned: true };
    },
    extendRowContext = (baseContext, rowSource, newRowIndex) => ({
      ...baseContext,
      rows: [
        ...baseContext.rows,
        {
          tableId: rowSource.tableId,
          index: newRowIndex,
          columns: rowSource.columns,
        },
      ],
      transitioned: false,
    }),
    dropFiltersForAll = (activeFilters, allTarget) => {
      if (allTarget.kind === "field") {
        const allTargetColumn = resolveColumn(allTarget);
        return stripFilterColumns(
          activeFilters,
          new Set([buildTableColumnKey(allTargetColumn)]),
        );
      }
      if (allTarget.kind !== "table")
        return throwDaxError(
          "DAX_ARGUMENT",
          "ALL accepts one physical table or qualified column.",
        );
      const affectedTableIds = new Set(
        tablesAffectedBy(findTable(allTarget.name).id).map(
          (affectedTable) => affectedTable.id,
        ),
      );
      return activeFilters.filter(
        (filterEntry) => !affectedTableIds.has(filterEntry.tableId),
      );
    },
    requireTableValue = (tableValue) =>
      typeof tableValue == "object" && tableValue !== null
        ? tableValue
        : throwDaxError("DAX_TYPE", "A model table expression is required."),
    evaluateMeasure = (measureDefinition, measureContext) => {
      if (
        measureStack.includes(measureDefinition.id) ||
        measureStack.length >= 48
      )
        return throwDaxError(
          "DAX_CYCLE",
          `Circular measure [${measureDefinition.name}].`,
        );
      measureStack.push(measureDefinition.id);
      try {
        const transitionedContext = transitionRowContext(measureContext);
        if (measureDefinition.operation === "dax")
          return requireScalarValue(
            evaluateNode(
              parsedExpressions.get(`measure:${measureDefinition.id}`),
              transitionedContext,
            ),
          );
        const visibleMeasureRows = visibleRows(
            measureDefinition.tableId,
            transitionedContext,
          ),
          measureValues = visibleMeasureRows.map((visibleRowIndex) =>
            measureDefinition.operation === "countRows"
              ? 1
              : readColumnValue(
                  {
                    tableId: measureDefinition.tableId,
                    column: measureDefinition.column,
                  },
                  visibleRowIndex,
                ),
          ),
          nonBlankValues = measureValues.filter(
            (measureValue) => measureValue !== null,
          );
        return measureDefinition.operation === "count" ||
          measureDefinition.operation === "countRows"
          ? nonBlankValues.length
          : measureDefinition.operation === "distinctCount"
            ? new Set(
                nonBlankValues.map((distinctValue) =>
                  JSON.stringify([typeof distinctValue, distinctValue]),
                ),
              ).size
            : nonBlankValues.length
              ? measureDefinition.operation === "sum"
                ? sumWithCompensation(nonBlankValues)
                : measureDefinition.operation === "average"
                  ? requireFiniteNumber(
                      sumWithCompensation(nonBlankValues) /
                        nonBlankValues.length,
                    )
                  : measureDefinition.operation === "min"
                    ? Math.min(...nonBlankValues)
                    : Math.max(...nonBlankValues)
              : null;
      } finally {
        measureStack.pop();
      }
    },
    readColumnValues = (columnNode, valueContext) => {
      const resolvedColumn = resolveColumn(columnNode);
      return {
        f: resolvedColumn,
        values: visibleRows(resolvedColumn.tableId, valueContext).map(
          (contextRowIndex) => readColumnValue(resolvedColumn, contextRowIndex),
        ),
      };
    };
  function evaluateNode(syntaxNode, rowContext) {
    if ((dataAccess.tick(), syntaxNode.kind === "literal"))
      return syntaxNode.value;
    if (syntaxNode.kind === "table") {
      const tableDefinition = findTable(syntaxNode.name);
      return {
        kind: "rows",
        tableId: tableDefinition.id,
        rows: visibleRows(tableDefinition.id, rowContext),
      };
    }
    if (syntaxNode.kind === "measure")
      return evaluateMeasure(findMeasure(syntaxNode.name), rowContext);
    if (syntaxNode.kind === "field") {
      const resolvedColumnRef = resolveColumn(syntaxNode),
        matchingRow = [...rowContext.rows]
          .reverse()
          .find(
            (contextRow) =>
              contextRow.tableId === resolvedColumnRef.tableId &&
              (!contextRow.columns ||
                contextRow.columns.includes(resolvedColumnRef.column)),
          );
      return matchingRow
        ? readColumnValue(resolvedColumnRef, matchingRow.index)
        : throwDaxError(
            "DAX_ROW",
            `No row context for ${syntaxNode.table}[${syntaxNode.column}]. Use an aggregate or SELECTEDVALUE in a measure.`,
          );
    }
    if (syntaxNode.kind === "unary") {
      const unaryOperand = requireScalarValue(
        evaluateNode(syntaxNode.value, rowContext),
      );
      return syntaxNode.op === "NOT"
        ? !requireLogicalOperand(unaryOperand)
        : syntaxNode.op === "+"
          ? unaryOperand
          : -requireNumericOperand(unaryOperand);
    }
    if (syntaxNode.kind === "binary") {
      const leftOperand = requireScalarValue(
        evaluateNode(syntaxNode.left, rowContext),
      );
      if (syntaxNode.op === "&&" && !requireLogicalOperand(leftOperand))
        return false;
      if (syntaxNode.op === "||" && requireLogicalOperand(leftOperand))
        return true;
      const rightOperand = requireScalarValue(
        evaluateNode(syntaxNode.right, rowContext),
      );
      if (syntaxNode.op === "&&" || syntaxNode.op === "||")
        return requireLogicalOperand(rightOperand);
      if (syntaxNode.op === "&")
        return `${leftOperand === null ? "" : typeof leftOperand == "boolean" ? String(leftOperand).toUpperCase() : leftOperand}${rightOperand === null ? "" : typeof rightOperand == "boolean" ? String(rightOperand).toUpperCase() : rightOperand}`;
      if (["=", "==", "<>", "<", ">", "<=", ">="].includes(syntaxNode.op)) {
        const comparisonResult = compareDaxValues(
          leftOperand,
          rightOperand,
          syntaxNode.op === "==",
        );
        return syntaxNode.op === "=" || syntaxNode.op === "=="
          ? comparisonResult === 0
          : syntaxNode.op === "<>"
            ? comparisonResult !== 0
            : syntaxNode.op === "<"
              ? comparisonResult !== undefined && comparisonResult < 0
              : syntaxNode.op === ">"
                ? comparisonResult !== undefined && comparisonResult > 0
                : syntaxNode.op === "<="
                  ? comparisonResult !== undefined && comparisonResult <= 0
                  : comparisonResult !== undefined && comparisonResult >= 0;
      }
      if (
        (syntaxNode.op === "*" &&
          (leftOperand === null || rightOperand === null)) ||
        (syntaxNode.op === "/" && leftOperand === null) ||
        (leftOperand === null &&
          rightOperand === null &&
          ["+", "-"].includes(syntaxNode.op))
      )
        return null;
      const leftNumber = requireNumericOperand(leftOperand),
        rightNumber = requireNumericOperand(rightOperand);
      return requireFiniteNumber(
        syntaxNode.op === "+"
          ? leftNumber + rightNumber
          : syntaxNode.op === "-"
            ? leftNumber - rightNumber
            : syntaxNode.op === "*"
              ? leftNumber * rightNumber
              : syntaxNode.op === "/"
                ? leftNumber / rightNumber
                : leftNumber ** rightNumber,
      );
    }
    const args = syntaxNode.args,
      evaluateArg = (argIndex) =>
        requireScalarValue(evaluateNode(args[argIndex], rowContext));
    switch (syntaxNode.name) {
      case "BLANK":
        return null;
      case "TRUE":
        return true;
      case "FALSE":
        return false;
      case "ISBLANK":
        return evaluateArg(0) === null;
      case "IF":
        return requireLogicalOperand(evaluateArg(0))
          ? evaluateArg(1)
          : args[2]
            ? evaluateArg(2)
            : null;
      case "COALESCE":
        for (
          let coalesceIndex = 0;
          coalesceIndex < args.length;
          coalesceIndex++
        ) {
          const coalesceValue = evaluateArg(coalesceIndex);
          if (coalesceValue !== null) return coalesceValue;
        }
        return null;
      case "DIVIDE": {
        if (args[2] && !isNumericLiteralArgument(args[2]))
          return throwDaxError(
            "DAX_ARGUMENT",
            "DIVIDE alternate result must be a numeric literal or BLANK().",
          );
        const dividend = evaluateArg(0),
          divisor = evaluateArg(1);
        return divisor === null || requireNumericOperand(divisor) === 0
          ? args[2]
            ? evaluateArg(2)
            : null
          : dividend === null
            ? null
            : requireFiniteNumber(
                requireNumericOperand(dividend) /
                  requireNumericOperand(divisor),
              );
      }
      case "SUM":
      case "AVERAGE":
      case "MIN":
      case "MAX": {
        const { f: valueColumnRef, values } = readColumnValues(
          args[0],
          rowContext,
        );
        if (
          tablesById
            .get(valueColumnRef.tableId)
            .columns.find((column) => column.name === valueColumnRef.column)
            .type !== "number"
        )
          return throwDaxError(
            "DAX_TYPE",
            `${syntaxNode.name} requires a number column in this subset.`,
          );
        const numericValues = values.filter(
          (value) => typeof value == "number",
        );
        return numericValues.length
          ? syntaxNode.name === "SUM"
            ? sumWithCompensation(numericValues)
            : syntaxNode.name === "AVERAGE"
              ? requireFiniteNumber(
                  sumWithCompensation(numericValues) / numericValues.length,
                )
              : syntaxNode.name === "MIN"
                ? Math.min(...numericValues)
                : Math.max(...numericValues)
          : null;
      }
      case "DISTINCTCOUNT": {
        const { values } = readColumnValues(args[0], rowContext);
        return values.length
          ? new Set(values.map(buildTypedValueKey)).size
          : null;
      }
      case "SELECTEDVALUE": {
        const { values } = readColumnValues(args[0], rowContext),
          uniqueValuesByKey = new Map(
            values.map((value) => [buildTypedValueKey(value), value]),
          );
        return uniqueValuesByKey.size === 1
          ? [...uniqueValuesByKey.values()][0]
          : args[1]
            ? evaluateArg(1)
            : null;
      }
      case "COUNTROWS":
        return (
          requireTableValue(evaluateNode(args[0], rowContext)).rows.length ||
          null
        );
      case "SUMX": {
        const iteratedTable = requireTableValue(
          evaluateNode(args[0], rowContext),
        );
        return sumWithCompensation(
          iteratedTable.rows.map((row) =>
            requireScalarValue(
              evaluateNode(
                args[1],
                extendRowContext(rowContext, iteratedTable, row),
              ),
            ),
          ),
        );
      }
      case "FILTER": {
        const filterSourceTable = requireTableValue(
          evaluateNode(args[0], rowContext),
        );
        return {
          ...filterSourceTable,
          rows: filterSourceTable.rows.filter((row) =>
            requireLogicalOperand(
              requireScalarValue(
                evaluateNode(
                  args[1],
                  extendRowContext(rowContext, filterSourceTable, row),
                ),
              ),
            ),
          ),
        };
      }
      case "ALL": {
        const allArgument = args[0];
        if (allArgument.kind === "table") {
          const tableId = findTable(allArgument.name).id;
          return {
            kind: "rows",
            tableId: tableId,
            rows: allRowIndexes(tableId),
          };
        }
        const allColumnRef = resolveColumn(allArgument),
          seenKeys = new Set(),
          rows = allRowIndexes(allColumnRef.tableId).filter(
            (distinctRowIndex) => {
              const typedValueKey = buildTypedValueKey(
                readColumnValue(allColumnRef, distinctRowIndex),
              );
              return seenKeys.has(typedValueKey)
                ? false
                : (seenKeys.add(typedValueKey), true);
            },
          );
        return {
          kind: "rows",
          tableId: allColumnRef.tableId,
          rows: rows,
          columns: [allColumnRef.column],
        };
      }
      case "RELATED": {
        const relatedColumnRef = resolveColumn(args[0]),
          innermostRow = rowContext.rows.at(-1);
        if (
          !innermostRow ||
          innermostRow.columns ||
          innermostRow.tableId === relatedColumnRef.tableId
        )
          return throwDaxError(
            "DAX_ROW",
            "RELATED requires a full row context on the related many side.",
          );
        const relatedRowIndex = dataAccess.parentRow(
          innermostRow.tableId,
          relatedColumnRef.tableId,
          innermostRow.index,
        );
        return relatedRowIndex === undefined
          ? throwDaxError(
              "DAX_RELATIONSHIP",
              "RELATED requires a unique many-to-one relationship path.",
            )
          : readColumnValue(relatedColumnRef, relatedRowIndex);
      }
      case "CALCULATE": {
        const filterSpecs = args.slice(1).map((filterArgument) => {
          if (filterArgument.kind === "call" && filterArgument.name === "ALL")
            return { remove: filterArgument.args[0] };
          if (
            filterArgument.kind === "table" ||
            (filterArgument.kind === "call" && filterArgument.name === "FILTER")
          ) {
            const tableResult = requireTableValue(
                evaluateNode(filterArgument, rowContext),
              ),
              relatedTables = tablesAffectedBy(tableResult.tableId).filter(
                (relatedTable) => relatedTable.id !== tableResult.tableId,
              );
            return !tableResult.columns &&
              relatedTables.some((leftTable) =>
                relatedTables.some(
                  (rightTable) =>
                    !dataAccess.canFilter(leftTable.id, rightTable.id) &&
                    !dataAccess.canFilter(rightTable.id, leftTable.id),
                ),
              )
              ? throwDaxError(
                  "DAX_FILTER",
                  "CALCULATE table filters spanning independent dimensions are outside this subset. Use explicit column predicates; FILTER remains available inside SUMX or COUNTROWS.",
                )
              : { filters: buildRowFilters(tableResult) };
          }
          const filterColumnRefs = [];
          if (
            (walkDaxSyntaxTree(filterArgument, (filterNode) => {
              (filterNode.kind === "field" &&
                filterColumnRefs.push(resolveColumn(filterNode)),
                (filterNode.kind === "measure" ||
                  filterNode.kind === "table" ||
                  (filterNode.kind === "call" &&
                    ![
                      "BLANK",
                      "TRUE",
                      "FALSE",
                      "ISBLANK",
                      "IF",
                      "COALESCE",
                      "DIVIDE",
                    ].includes(filterNode.name))) &&
                  throwDaxError(
                    "DAX_FILTER",
                    "Boolean CALCULATE filters support same-table columns, scalar functions and literals; use FILTER for a table expression.",
                  ));
            }),
            !filterColumnRefs.length ||
              new Set(
                filterColumnRefs.map((filterTableId) => filterTableId.tableId),
              ).size !== 1)
          )
            return throwDaxError(
              "DAX_FILTER",
              "A Boolean CALCULATE filter must reference columns from one table.",
            );
          const tableId = filterColumnRefs[0].tableId,
            matchingRows = allRowIndexes(tableId).filter((filterCandidateRow) =>
              requireLogicalOperand(
                requireScalarValue(
                  evaluateNode(filterArgument, {
                    ...rowContext,
                    rows: [
                      ...rowContext.rows,
                      { tableId: tableId, index: filterCandidateRow },
                    ],
                    transitioned: false,
                  }),
                ),
              ),
            );
          return {
            filters: [
              createRowSetFilter(
                tableId,
                [
                  ...new Set(
                    filterColumnRefs.map(
                      (filterColumnName) => filterColumnName.column,
                    ),
                  ),
                ],
                matchingRows,
              ),
            ],
          };
        });
        let filters = [...transitionRowContext(rowContext).filters];
        for (const filterSpec of filterSpecs)
          filterSpec.remove &&
            (filters = dropFiltersForAll(filters, filterSpec.remove));
        return (
          (filters = mergeFilterSets(
            filters,
            filterSpecs.flatMap((spec) => spec.filters ?? []),
          )),
          evaluateNode(args[0], {
            ...rowContext,
            filters: filters,
            transitioned: true,
          })
        );
      }
      case "TOTALYTD": {
        const dateColumnRef = resolveColumn(args[1]),
          calendarTable = tablesById.get(dateColumnRef.tableId);
        if (calendarTable.dateColumn !== dateColumnRef.column)
          return throwDaxError(
            "DAX_CALENDAR",
            "TOTALYTD requires a marked physical calendar date column.",
          );
        const calendarDates = allRowIndexes(calendarTable.id).map(
          (calendarRowIndex) =>
            readColumnValue(dateColumnRef, calendarRowIndex),
        );
        validateMarkedCalendarDays(calendarDates);
        const transitionedState = transitionRowContext(rowContext),
          yearDateStrings = visibleRows(calendarTable.id, transitionedState)
            .map((yearRowIndex) =>
              String(readColumnValue(dateColumnRef, yearRowIndex)),
            )
            .sort();
        if (!yearDateStrings.length) return null;
        const latestDate = yearDateStrings.at(-1),
          yearStartDate = latestDate.slice(0, 4) + "-01-01",
          filters = transitionedState.filters.filter(
            (filter) => filter.tableId !== calendarTable.id,
          );
        return (
          filters.push({
            tableId: calendarTable.id,
            columns: calendarTable.columns.map((column) => column.name),
            test: (dateRowIndex) => {
              const rowDate = String(
                readColumnValue(dateColumnRef, dateRowIndex),
              );
              return rowDate >= yearStartDate && rowDate <= latestDate;
            },
          }),
          evaluateNode(args[0], {
            ...transitionedState,
            filters: filters,
            transitioned: true,
          })
        );
      }
    }
  }
  const collectMeasureTables = (
    targetMeasure,
    visitedMeasureIds = new Set(),
  ) => {
    if (visitedMeasureIds.has(targetMeasure.id)) return new Set();
    visitedMeasureIds.add(targetMeasure.id);
    const referencedTableIds = new Set([targetMeasure.tableId]);
    return (
      targetMeasure.operation === "dax" &&
        walkDaxSyntaxTree(
          parsedExpressions.get(`measure:${targetMeasure.id}`),
          (daxNode) => {
            if (
              (daxNode.kind === "call" &&
                [
                  "SUM",
                  "AVERAGE",
                  "MIN",
                  "MAX",
                  "DISTINCTCOUNT",
                  "SELECTEDVALUE",
                ].includes(daxNode.name) &&
                referencedTableIds.add(resolveColumn(daxNode.args[0]).tableId),
              daxNode.kind === "call" &&
                ["SUMX", "COUNTROWS"].includes(daxNode.name))
            ) {
              let aggregateSource = daxNode.args[0];
              for (
                ;
                aggregateSource.kind === "call" &&
                ["FILTER", "ALL"].includes(aggregateSource.name);
              )
                aggregateSource = aggregateSource.args[0];
              aggregateSource.kind === "table"
                ? referencedTableIds.add(findTable(aggregateSource.name).id)
                : aggregateSource.kind === "field" &&
                  referencedTableIds.add(
                    resolveColumn(aggregateSource).tableId,
                  );
            }
            daxNode.kind === "measure" &&
              collectMeasureTables(
                findMeasure(daxNode.name),
                visitedMeasureIds,
              ).forEach((nestedTableId) =>
                referencedTableIds.add(nestedTableId),
              );
          },
        ),
      referencedTableIds
    );
  };
  return {
    calculatedColumns: () =>
      definition.tables.flatMap((table) =>
        table.columns
          .filter((column) => column.expression !== undefined)
          .map((calculatedColumn) => ({
            tableId: table.id,
            name: calculatedColumn.name,
            type: calculatedColumn.type,
            values: allRowIndexes(table.id).map((calculatedRowIndex) =>
              readColumnValue(
                { tableId: table.id, column: calculatedColumn.name },
                calculatedRowIndex,
              ),
            ),
          })),
      ),
    query(queryRequest, queryResult) {
      const selectedMeasures = queryRequest.measureIds.map((measureId) =>
          definition.measures.find((measure) => measure.id === measureId),
        ),
        selectedTableIds = new Set(
          selectedMeasures.flatMap((selectedMeasure) => [
            ...collectMeasureTables(selectedMeasure),
          ]),
        ),
        filters = queryRequest.filters.map((filter) => {
          const column = tablesById
            .get(filter.tableId)
            .columns.find((column) => column.name === filter.column);
          return (
            !["blank", "notBlank"].includes(filter.operator) &&
              (filter.value === undefined ||
                !isValidColumnValue(filter.value, {
                  ...column,
                  nullable: true,
                })) &&
              throwDaxError(
                "DAX_FILTER",
                "Filter value does not match its declared column type.",
              ),
            ["gt", "gte", "lt", "lte"].includes(filter.operator) &&
              (!["number", "date"].includes(column.type) ||
                filter.value == null) &&
              throwDaxError(
                "DAX_FILTER",
                "Ordered query filters require a nonblank number or date.",
              ),
            selectedMeasures.length &&
              ![...selectedTableIds].some((measureTableId) =>
                dataAccess.canFilter(filter.tableId, measureTableId),
              ) &&
              throwDaxError(
                "DAX_FILTER",
                "A query filter cannot reach any selected measure.",
              ),
            {
              tableId: filter.tableId,
              columns: [filter.column],
              test: (queryRowIndex) =>
                matchesQueryFilter(
                  readColumnValue(
                    { tableId: filter.tableId, column: filter.column },
                    queryRowIndex,
                  ),
                  filter.operator,
                  filter.value,
                ),
            }
          );
        }),
        queryContext = { filters: filters, rows: [] };
      for (const table of queryResult.tables)
        ((table.included = visibleRows(table.id, queryContext).length),
          (table.excluded = table.rows - table.included));
      for (const groupItem of queryRequest.groupBy)
        for (const selectedTableId of selectedTableIds)
          dataAccess.canFilter(groupItem.tableId, selectedTableId) ||
            throwDaxError(
              "DAX_FILTER",
              `Group ${groupItem.column} cannot filter every table referenced by the selected measures.`,
            );
      ((queryResult.headers = [
        ...queryRequest.groupBy.map(
          (groupEntry) =>
            `${tablesById.get(groupEntry.tableId).name}.${groupEntry.column}`,
        ),
        ...selectedMeasures.map((headerMeasure) => headerMeasure.name),
      ]),
        new Set(queryResult.headers.map(caseFoldText)).size !==
          queryResult.headers.length &&
          throwDaxError(
            "DAX_REFERENCE",
            "Group and measure headings must be distinct.",
          ));
      const requireNumericTotal = (totalValue) =>
        totalValue === null || typeof totalValue == "number"
          ? totalValue
          : throwDaxError(
              "DAX_TYPE",
              "Published model measures must return numbers or BLANK. Use typed calculated columns for text or boolean results.",
            );
      if (
        ((queryResult.totals = selectedMeasures.map((totalMeasure) =>
          requireNumericTotal(evaluateMeasure(totalMeasure, queryContext)),
        )),
        !queryRequest.groupBy.length)
      ) {
        queryResult.rows = [{ keys: [], values: [...queryResult.totals] }];
        return;
      }
      const groupsByKey = new Map();
      for (const groupTableId of new Set(
        selectedMeasures.map((measureForTable) => measureForTable.tableId),
      ))
        for (const groupRowIndex of visibleRows(groupTableId, queryContext)) {
          dataAccess.tick();
          const groupKeyValues = queryRequest.groupBy.map((groupDefinition) => {
            const groupParentRowIndex = dataAccess.parentRow(
              groupTableId,
              groupDefinition.tableId,
              groupRowIndex,
            );
            return groupParentRowIndex === undefined
              ? throwDaxError(
                  "DAX_RELATIONSHIP",
                  "A group has no relationship path.",
                )
              : readColumnValue(groupDefinition, groupParentRowIndex);
          });
          (groupsByKey.set(JSON.stringify(groupKeyValues), groupKeyValues),
            groupsByKey.size > dataModelLimits.groups &&
              throwDaxError(
                "DAX_LIMIT",
                "The model query exceeds 1,000 groups.",
              ));
        }
      queryResult.rows = [...groupsByKey.values()].map((keys) => {
        const groupContext = {
          rows: [],
          filters: [
            ...filters,
            ...queryRequest.groupBy.map((groupItemForFilter, groupIndex) => ({
              tableId: groupItemForFilter.tableId,
              columns: [groupItemForFilter.column],
              test: (filteredRowIndex) =>
                readColumnValue(groupItemForFilter, filteredRowIndex) ===
                keys[groupIndex],
            })),
          ],
        };
        return {
          keys: keys,
          values: selectedMeasures.map((valueMeasure) =>
            requireNumericTotal(evaluateMeasure(valueMeasure, groupContext)),
          ),
        };
      });
    },
  };
}
function isValidColumnValue(value, column) {
  return value === null
    ? column.nullable
    : column.type === "number"
      ? typeof value == "number" && Number.isFinite(value)
      : column.type === "boolean"
        ? typeof value == "boolean"
        : column.type === "text"
          ? typeof value == "string" && value.length <= 32767
          : typeof value == "string" &&
            /^\d{4}-\d{2}-\d{2}$/.test(value) &&
            Number(value.slice(0, 4)) >= 1 &&
            Number.isFinite(new Date(value + "T00:00:00Z").getTime()) &&
            new Date(value + "T00:00:00Z").toISOString().slice(0, 10) === value;
}
function validateMarkedCalendarDays(days) {
  const sortedDays = days.map((day) => String(day)).sort(),
    year = sortedDays[0]?.slice(0, 4);
  (!sortedDays.length ||
    new Set(sortedDays).size !== sortedDays.length ||
    sortedDays[0] !== `${year}-01-01` ||
    sortedDays.at(-1) !== `${year}-12-31`) &&
    throwDaxError(
      "DAX_CALENDAR",
      "A marked calendar must contain every unique day of one complete calendar year, without blanks.",
    );
  for (let dayIndex = 1; dayIndex < sortedDays.length; dayIndex++)
    new Date(sortedDays[dayIndex] + "T00:00:00Z").getTime() -
      new Date(sortedDays[dayIndex - 1] + "T00:00:00Z").getTime() !==
      864e5 &&
      throwDaxError(
        "DAX_CALENDAR",
        "The marked calendar has a missing or invalid day.",
      );
}
function matchesQueryFilter(value, operator, filterValue) {
  switch (operator) {
    case "blank":
      return value === null || value === "";
    case "notBlank":
      return value !== null && value !== "";
    case "equals":
      return value === filterValue;
    case "notEquals":
      return value !== filterValue;
    case "gt":
      return value !== null && value > filterValue;
    case "gte":
      return value !== null && value >= filterValue;
    case "lt":
      return value !== null && value < filterValue;
    case "lte":
      return value !== null && value <= filterValue;
    default:
      throw new DaxError("DAX_FILTER", "Unsupported query filter.");
  }
}
class DaxModelError extends Error {
  constructor(code, message, context = {}) {
    (super(message), (this.code = code), (this.context = context));
  }
}
const throwDaxModelError = (code, message, context = {}) => {
    throw new DaxModelError(code, message, context);
  },
  typedValueKey = (value) => JSON.stringify([typeof value, value]);
function isValidColumnValueForType(value, column) {
  if (value === null) return column.nullable;
  if (column.type === "number")
    return typeof value == "number" && Number.isFinite(value);
  if (column.type === "boolean") return typeof value == "boolean";
  if (column.type === "text") return typeof value == "string";
  if (
    typeof value != "string" ||
    !/^\d{4}-\d{2}-\d{2}$/.test(value) ||
    Number(value.slice(0, 4)) < 1
  )
    return false;
  const parsedDate = new Date(value + "T00:00:00Z");
  return (
    !isNaN(parsedDate.getTime()) &&
    parsedDate.toISOString().slice(0, 10) === value
  );
}
const decodeColumnValue = (column, rowIndex) =>
  column.nulls[rowIndex]
    ? null
    : column.dictionary
      ? column.dictionary[column.data[rowIndex]]
      : column.definition.type === "boolean"
        ? !!column.data[rowIndex]
        : column.data[rowIndex];
function encodeColumnDictionary(definition, values) {
  const nulls = new Uint8Array(values.length),
    dictionary = [],
    dictionaryIndexByText = new Map(),
    encodedData =
      definition.type === "number"
        ? new Float64Array(values.length)
        : definition.type === "boolean"
          ? new Uint8Array(values.length)
          : new Uint32Array(values.length);
  return (
    values.forEach((value, valueIndex) => {
      if (value === null) nulls[valueIndex] = 1;
      else if (definition.type === "number") encodedData[valueIndex] = value;
      else if (definition.type === "boolean")
        encodedData[valueIndex] = value ? 1 : 0;
      else {
        const text = String(value);
        (dictionaryIndexByText.has(text) ||
          (dictionaryIndexByText.set(text, dictionary.length),
          dictionary.push(text)),
          (encodedData[valueIndex] = dictionaryIndexByText.get(text)));
      }
    }),
    {
      definition: definition,
      nulls: nulls,
      data: encodedData,
      ...(["text", "date"].includes(definition.type)
        ? { dictionary: dictionary }
        : {}),
    }
  );
}
function resolveModelTableRange(workbook, sourceTable) {
  const sheet = workbook.sheets.find(
    (sheet) => sheet.id === sourceTable.sheetId,
  );
  if (!sheet)
    return throwDaxModelError(
      "SOURCE",
      `Model table ${sourceTable.name} has a missing worksheet.`,
      {
        tableId: sourceTable.id,
      },
    );
  if (!sourceTable.tableId) return sourceTable.range;
  const structuredTable = sheet.tables?.find(
    (table) => table.id === sourceTable.tableId,
  );
  if (!structuredTable)
    return throwDaxModelError(
      "SOURCE",
      `The structured table for ${sourceTable.name} is missing.`,
      {
        tableId: sourceTable.id,
      },
    );
  const rangeAddress = parseRangeAddress(structuredTable.range);
  return `${columnIndexToLetters(rangeAddress.firstCol)}${rangeAddress.firstRow + 1}:${columnIndexToLetters(rangeAddress.lastCol)}${rangeAddress.lastRow + 1 - (structuredTable.showTotals ? 1 : 0)}`;
}
function buildModelTableFromRange(
  workbook,
  sheetId,
  range,
  tableId,
  options = {},
) {
  const sheet = workbook.sheets.find((sheet) => sheet.id === sheetId);
  if (!sheet) throw Error("Choose a source worksheet.");
  const tableDefinition = {
      id: crypto.randomUUID(),
      name: sheet.name,
      sheetId: sheetId,
      range: range ?? formatOccupiedRangeAddress(sheet),
      ...(tableId ? { tableId: tableId } : {}),
      columns: [],
    },
    snapshot = readSourceHeaders(
      workbook,
      [
        {
          sheetId: sheetId,
          range: resolveModelTableRange(workbook, tableDefinition),
        },
      ],
      options,
    ).snapshots[0];
  return (
    (tableDefinition.range = snapshot.source.range),
    (tableDefinition.columns = snapshot.headers.map((header, columnIndex) => {
      const columnValues = snapshot.records.map(
          (record) => record.values[columnIndex],
        ),
        nonNullValues = columnValues.filter((value) => value !== null);
      if (new Set(nonNullValues.map((value) => typeof value)).size > 1)
        throw Error(
          `Column ${header} contains mixed types. Clean the source before adding it to the model.`,
        );
      return {
        name: header,
        type:
          nonNullValues.length && typeof nonNullValues[0] == "number"
            ? "number"
            : nonNullValues.length && typeof nonNullValues[0] == "boolean"
              ? "boolean"
              : "text",
        nullable: columnValues.some((value) => value === null),
      };
    })),
    tableDefinition
  );
}
function normalizeModelHostPermissions(workbook, hostPermissions) {
  if (
    hostPermissions.readOnly !== undefined &&
    typeof hostPermissions.readOnly != "boolean"
  )
    throw Error("Host read-only must be boolean.");
  return {
    readOnly: !!hostPermissions.readOnly,
    includeHidden: !!hostPermissions.includeHidden,
    allowedSheetIds: [
      ...new Set(
        hostPermissions.allowedSheetIds ??
          workbook.sheets.map((sheet) => sheet.id),
      ),
    ].sort(),
  };
}
function buildModelTableSnapshots(
  workbook,
  modelDefinition,
  hostPermissions,
  result,
) {
  normalizeModelHostPermissions(workbook, hostPermissions);
  for (const table of modelDefinition.tables)
    workbook.sheets.find((sheet) => sheet.id === table.sheetId)
      ?.dataModelReport &&
      throwDaxModelError(
        "SNAPSHOT_SOURCE",
        "Model report snapshots cannot be calculation sources. Use the original typed source tables; retained measure caches are not trusted inputs.",
      );
  const sourceRanges = modelDefinition.tables.map((table) => ({
    sheetId: table.sheetId,
    range: resolveModelTableRange(workbook, table),
  }));
  let totalCells = 0;
  for (const sourceRange of sourceRanges) {
    const rangeAddress = parseRangeAddress(sourceRange.range),
      cellCount =
        (rangeAddress.lastRow - rangeAddress.firstRow + 1) *
        (rangeAddress.lastCol - rangeAddress.firstCol + 1);
    ((totalCells += cellCount),
      (cellCount > dataModelLimits.cellsPerTable ||
        rangeAddress.lastRow - rangeAddress.firstRow >
          dataModelLimits.rowsPerTable ||
        rangeAddress.lastCol - rangeAddress.firstCol + 1 >
          dataModelLimits.columns) &&
        throwDaxModelError(
          "LIMIT",
          "A model table supports at most 500 data rows, 50 fields and 2,000 cells including headers.",
        ));
  }
  totalCells > dataModelLimits.modelCells &&
    throwDaxModelError("LIMIT", "The data model exceeds 16,000 source cells.");
  const sourceHeaders = readSourceHeaders(
      workbook,
      sourceRanges,
      hostPermissions,
    ),
    tableSnapshots = new Map();
  return (
    workbook.calculation?.mode === "manual" &&
      sourceHeaders.dependencyCells.some(
        (dependencyCell) =>
          typeof dependencyCell.cell?.raw == "string" &&
          dependencyCell.cell.raw.startsWith("=") &&
          !dependencyCell.cell.literal,
      ) &&
      throwDaxModelError(
        "MANUAL",
        "Model source formulas require Automatic calculation. In Manual mode use an explicitly reviewed values snapshot; pending or historical formula values cannot enter the model.",
      ),
    sourceHeaders.snapshots.forEach((snapshot, tableIndex) => {
      const definition = modelDefinition.tables[tableIndex],
        storedColumns = definition.columns.filter(
          (column) => column.expression === undefined,
        );
      (snapshot.headers.length !== storedColumns.length ||
        snapshot.headers.some(
          (header, headerIndex) => header !== storedColumns[headerIndex].name,
        )) &&
        throwDaxModelError(
          "SCHEMA",
          `The source schema of ${definition.name} changed. Review the model columns.`,
          { tableId: definition.id },
        );
      const columns = new Map();
      let bytes = 0;
      (storedColumns.forEach((storedColumn, columnIndex) => {
        const values = snapshot.records.map((record) => {
            const cellValue = record.values[columnIndex];
            return (
              isValidColumnValueForType(cellValue, storedColumn) ||
                throwDaxModelError(
                  "TYPE",
                  `${definition.name}.${storedColumn.name} has an invalid ${storedColumn.type} value on worksheet row ${record.row}.`,
                  {
                    tableId: definition.id,
                    column: storedColumn.name,
                    row: record.row,
                  },
                ),
              cellValue
            );
          }),
          encodedColumn = encodeColumnDictionary(storedColumn, values);
        (columns.set(storedColumn.name, encodedColumn),
          (bytes +=
            encodedColumn.data.byteLength +
            encodedColumn.nulls.byteLength +
            (encodedColumn.dictionary?.reduce(
              (total, dictionaryEntry) => total + dictionaryEntry.length * 2,
              0,
            ) ?? 0)));
      }),
        tableSnapshots.set(definition.id, {
          definition: definition,
          columns: columns,
          count: snapshot.records.length,
          range: snapshot.source.range,
          sheetName: snapshot.sheetName,
          bytes: bytes,
        }));
    }),
    (result.sourceHash = fingerprintCanonicalValue({
      sources: sourceHeaders.snapshots,
      dependencies: sourceHeaders.dependencyCells,
      datePolicy: workbook.datePolicy,
    })),
    tableSnapshots
  );
}
function evaluateDataModelQuery(workbook, definition, query, options = {}) {
  const result = {
    complete: false,
    diagnostics: [],
    tables: [],
    rows: [],
    totals: [],
    headers: [],
    sourceHash: "",
    operations: 0,
    engine: columnarModelFormatVersion,
  };
  try {
    const model = validateDataModel(definition),
      modelQuery = validateModelQuery(query, model),
      modelTableSnapshots = buildModelTableSnapshots(
        workbook,
        model,
        options,
        result,
      ),
      tableIds = model.tables.map((table) => table.id),
      tick = (amount = 1) => {
        ((result.operations += amount),
          result.operations > dataModelLimits.operations &&
            throwDaxModelError(
              "LIMIT",
              "Data model execution exceeded 2,000,000 operations. Reduce the model or query.",
            ));
      },
      relationshipsByOneTable = new Map(
        tableIds.map((tableId) => [
          tableId,
          model.relationships.filter(
            (relationship) => relationship.one.tableId === tableId,
          ),
        ]),
      ),
      incomingCounts = new Map(
        tableIds.map((countedTableId) => [countedTableId, 0]),
      );
    for (const relationship of model.relationships)
      incomingCounts.set(
        relationship.many.tableId,
        incomingCounts.get(relationship.many.tableId) + 1,
      );
    const order = tableIds.filter(
      (rootTableId) => !incomingCounts.get(rootTableId),
    );
    for (let orderIndex = 0; orderIndex < order.length; orderIndex++)
      for (const relationship of relationshipsByOneTable.get(
        order[orderIndex],
      )) {
        const remainingCount =
          incomingCounts.get(relationship.many.tableId) - 1;
        (incomingCounts.set(relationship.many.tableId, remainingCount),
          remainingCount || order.push(relationship.many.tableId));
      }
    order.length !== tableIds.length &&
      throwDaxModelError(
        "CYCLE",
        "Relationships form a cycle. Use a single-direction acyclic model.",
      );
    const filterPathsByTable = new Map();
    for (const sourceTableId of tableIds) {
      const relationshipPaths = new Map([[sourceTableId, []]]);
      for (const pathTableId of order) {
        const currentPath = relationshipPaths.get(pathTableId);
        if (currentPath)
          for (const relationship of relationshipsByOneTable.get(pathTableId))
            (relationshipPaths.has(relationship.many.tableId) &&
              throwDaxModelError(
                "AMBIGUOUS",
                `More than one filter path connects ${modelTableSnapshots.get(sourceTableId).definition.name} to ${modelTableSnapshots.get(relationship.many.tableId).definition.name}. Remove the ambiguous relationship.`,
              ),
              relationshipPaths.set(relationship.many.tableId, [
                ...currentPath,
                relationship.id,
              ]));
      }
      filterPathsByTable.set(sourceTableId, relationshipPaths);
    }
    const parentRowIndexes = new Map();
    for (const relationship of model.relationships) {
      const oneSnapshot = modelTableSnapshots.get(relationship.one.tableId),
        manySnapshot = modelTableSnapshots.get(relationship.many.tableId),
        oneRowByKey = new Map(),
        seenKeys = new Set(),
        parentRows = new Int32Array(manySnapshot.count).fill(-1);
      for (
        let oneRowIndex = 0;
        oneRowIndex < oneSnapshot.count;
        oneRowIndex++
      ) {
        tick();
        const oneKeyValue = decodeColumnValue(
            oneSnapshot.columns.get(relationship.one.column),
            oneRowIndex,
          ),
          oneTypedKey = typedValueKey(oneKeyValue);
        ((oneKeyValue === null || oneKeyValue === "") &&
          throwDaxModelError(
            "BLANK_KEY",
            `The one side of ${relationship.id} contains a blank key.`,
            {
              tableId: relationship.one.tableId,
              column: relationship.one.column,
              row:
                parseRangeAddress(oneSnapshot.range).firstRow + oneRowIndex + 2,
              relationshipId: relationship.id,
            },
          ),
          oneRowByKey.has(oneTypedKey) &&
            throwDaxModelError(
              "DUPLICATE_KEY",
              `The one side of ${relationship.id} contains a duplicate key.`,
              {
                tableId: relationship.one.tableId,
                column: relationship.one.column,
                row:
                  parseRangeAddress(oneSnapshot.range).firstRow +
                  oneRowIndex +
                  2,
                relationshipId: relationship.id,
              },
            ),
          oneRowByKey.set(oneTypedKey, oneRowIndex));
      }
      let missingKeyCount = 0;
      for (
        let manyRowIndex = 0;
        manyRowIndex < manySnapshot.count;
        manyRowIndex++
      ) {
        tick();
        const manyKeyValue = decodeColumnValue(
            manySnapshot.columns.get(relationship.many.column),
            manyRowIndex,
          ),
          manyTypedKey = typedValueKey(manyKeyValue),
          matchedOneRow = oneRowByKey.get(manyTypedKey);
        (relationship.cardinality === "one-to-one" &&
          seenKeys.has(manyTypedKey) &&
          throwDaxModelError(
            "CARDINALITY",
            `The related side of ${relationship.id} is not unique. Choose one-to-many or correct its keys.`,
            { relationshipId: relationship.id },
          ),
          seenKeys.add(manyTypedKey),
          matchedOneRow === undefined
            ? (missingKeyCount++,
              result.diagnostics.length < 100 &&
                result.diagnostics.push({
                  code: "MISSING_KEY",
                  message: `${manySnapshot.definition.name}.${relationship.many.column} has no dimension key; correct the source before calculating.`,
                  tableId: relationship.many.tableId,
                  column: relationship.many.column,
                  row:
                    parseRangeAddress(manySnapshot.range).firstRow +
                    manyRowIndex +
                    2,
                  relationshipId: relationship.id,
                }))
            : (parentRows[manyRowIndex] = matchedOneRow));
      }
      (missingKeyCount > 100 &&
        result.diagnostics.push({
          code: "MISSING_KEY_LIMIT",
          message: `${relationship.id}: ${missingKeyCount} missing dimension keys; only the first 100 diagnostics are shown.`,
          relationshipId: relationship.id,
        }),
        parentRowIndexes.set(relationship.id, parentRows));
    }
    const includedRowMasks = new Map(
      [...modelTableSnapshots].map(([snapshotTableId, tableSnapshot]) => [
        snapshotTableId,
        new Uint8Array(tableSnapshot.count).fill(1),
      ]),
    );
    if (
      ((result.tables = [...modelTableSnapshots.values()].map((snapshot) => ({
        id: snapshot.definition.id,
        name: snapshot.definition.name,
        sheetId: snapshot.definition.sheetId,
        range: snapshot.range,
        rows: snapshot.count,
        included: snapshot.count,
        excluded: 0,
        columnarBytes: snapshot.bytes,
      }))),
      result.diagnostics.length)
    )
      return result;
    if (
      model.measures.some((measure) => measure.operation === "dax") ||
      model.tables.some((table) =>
        table.columns.some((column) => column.expression !== undefined),
      )
    ) {
      const daxEvaluator = evaluateDaxExpression({
        definition: model,
        count: (countTableId) => modelTableSnapshots.get(countTableId).count,
        value: (valueTableId, columnName, valueRowIndex) =>
          decodeColumnValue(
            modelTableSnapshots.get(valueTableId).columns.get(columnName),
            valueRowIndex,
          ),
        order: order,
        tick: tick,
        canFilter: (fromTableId, toTableId) =>
          filterPathsByTable.get(fromTableId)?.has(toTableId) ?? false,
        parentRow: (childTableId, parentTableId, childRowIndex) => {
          const pathToChild = filterPathsByTable
            .get(parentTableId)
            ?.get(childTableId);
          if (!pathToChild) return;
          let parentRowIndex = childRowIndex;
          for (
            let pathIndex = pathToChild.length - 1;
            pathIndex >= 0;
            pathIndex--
          )
            (tick(),
              (parentRowIndex = parentRowIndexes.get(pathToChild[pathIndex])[
                parentRowIndex
              ]));
          return parentRowIndex;
        },
      });
      result.calculatedColumns = daxEvaluator.calculatedColumns();
      for (const calculatedColumn of result.calculatedColumns) {
        const calculatedTableSnapshot = modelTableSnapshots.get(
            calculatedColumn.tableId,
          ),
          encodedColumn = encodeColumnDictionary(
            calculatedTableSnapshot.definition.columns.find(
              (column) => column.name === calculatedColumn.name,
            ),
            calculatedColumn.values,
          );
        (calculatedTableSnapshot.columns.set(
          calculatedColumn.name,
          encodedColumn,
        ),
          (result.tables.find(
            (table) => table.id === calculatedColumn.tableId,
          ).columnarBytes +=
            encodedColumn.data.byteLength +
            encodedColumn.nulls.byteLength +
            (encodedColumn.dictionary?.reduce(
              (total, dictionaryEntry) => total + dictionaryEntry.length * 2,
              0,
            ) ?? 0)));
      }
      return (
        daxEvaluator.query(modelQuery, result),
        (result.rows.length + 2) * result.headers.length >
          dataModelLimits.outputCells &&
          throwDaxModelError(
            "LIMIT",
            "The model result exceeds 20,000 output cells.",
          ),
        (result.complete = true),
        result
      );
    }
    const selectedMeasures = modelQuery.measureIds.map((measureId) =>
      model.measures.find((measure) => measure.id === measureId),
    );
    for (const filter of modelQuery.filters) {
      const filterTableSnapshot = modelTableSnapshots.get(filter.tableId),
        filterColumnData = filterTableSnapshot.columns.get(filter.column),
        filterRowMask = includedRowMasks.get(filter.tableId);
      (!["blank", "notBlank"].includes(filter.operator) &&
        (filter.value === undefined ||
          !isValidColumnValueForType(filter.value, {
            ...filterColumnData.definition,
            nullable: true,
          })) &&
        throwDaxModelError(
          "FILTER",
          "Filter values must match the declared column type.",
        ),
        ["gt", "gte", "lt", "lte"].includes(filter.operator) &&
          (!["number", "date"].includes(filterColumnData.definition.type) ||
            filter.value === null) &&
          throwDaxModelError(
            "FILTER",
            "Ordered filters require a nonblank number or ISO date.",
          ),
        selectedMeasures.length &&
          !selectedMeasures.some((selectedMeasure) =>
            filterPathsByTable.get(filter.tableId).has(selectedMeasure.tableId),
          ) &&
          throwDaxModelError(
            "DISCONNECTED",
            "A filter cannot reach any selected measure. Connect its table or remove the filter.",
          ));
      for (let rowIndex = 0; rowIndex < filterTableSnapshot.count; rowIndex++) {
        tick();
        const cellValue = decodeColumnValue(filterColumnData, rowIndex);
        matchesComparisonOperator(cellValue, filter) ||
          (filterRowMask[rowIndex] = 0);
      }
    }
    for (const orderedTableId of order)
      for (const relationship of relationshipsByOneTable.get(orderedTableId)) {
        const parentRowMask = includedRowMasks.get(orderedTableId),
          childRowMask = includedRowMasks.get(relationship.many.tableId),
          parentRows = parentRowIndexes.get(relationship.id);
        for (
          let childRowIndex = 0;
          childRowIndex < childRowMask.length;
          childRowIndex++
        )
          (tick(),
            parentRowMask[parentRows[childRowIndex]] ||
              (childRowMask[childRowIndex] = 0));
      }
    for (const table of result.tables)
      ((table.included = includedRowMasks
        .get(table.id)
        .reduce((total, isIncluded) => total + isIncluded, 0)),
        (table.excluded = table.rows - table.included));
    for (const groupField of modelQuery.groupBy)
      for (const selectedMeasure of selectedMeasures)
        filterPathsByTable
          .get(groupField.tableId)
          .has(selectedMeasure.tableId) ||
          throwDaxModelError(
            "DISCONNECTED",
            `Group field ${groupField.column} cannot filter measure ${selectedMeasure.name}. Use dimensions shared by every selected fact table.`,
          );
    ((result.headers = [
      ...modelQuery.groupBy.map(
        (groupField) =>
          `${modelTableSnapshots.get(groupField.tableId).definition.name}.${groupField.column}`,
      ),
      ...selectedMeasures.map((selectedMeasure) => selectedMeasure.name),
    ]),
      new Set(result.headers.map((header) => header.toLowerCase())).size !==
        result.headers.length &&
        throwDaxModelError(
          "SCHEMA",
          "Group and measure headings must be distinct.",
        ));
    const readGroupKeys = (sourceTableId, rowIndex) =>
        modelQuery.groupBy.map((groupField) => {
          let groupRowIndex = rowIndex;
          const groupPath = filterPathsByTable
            .get(groupField.tableId)
            .get(sourceTableId);
          for (
            let pathIndex = groupPath.length - 1;
            pathIndex >= 0;
            pathIndex--
          )
            (tick(),
              (groupRowIndex = parentRowIndexes.get(groupPath[pathIndex])[
                groupRowIndex
              ]));
          return decodeColumnValue(
            modelTableSnapshots
              .get(groupField.tableId)
              .columns.get(groupField.column),
            groupRowIndex,
          );
        }),
      groupBuckets = new Map(),
      totalAccumulators = selectedMeasures.map(() => createSumAccumulator());
    for (const [measureIndex, selectedMeasure] of selectedMeasures.entries()) {
      const measureSnapshot = modelTableSnapshots.get(selectedMeasure.tableId),
        includedRowMask = includedRowMasks.get(selectedMeasure.tableId),
        columnData = selectedMeasure.column
          ? measureSnapshot.columns.get(selectedMeasure.column)
          : undefined;
      for (let rowIndex = 0; rowIndex < measureSnapshot.count; rowIndex++) {
        if ((tick(), !includedRowMask[rowIndex])) continue;
        const aggregateValue =
          selectedMeasure.operation === "countRows"
            ? 1
            : decodeColumnValue(columnData, rowIndex);
        if (
          (addAggregateValue(totalAccumulators[measureIndex], aggregateValue),
          !modelQuery.groupBy.length)
        )
          continue;
        const keys = readGroupKeys(selectedMeasure.tableId, rowIndex),
          groupKey = JSON.stringify(keys);
        (groupBuckets.has(groupKey) ||
          (groupBuckets.size >= dataModelLimits.groups &&
            throwDaxModelError(
              "LIMIT",
              "The model query exceeds 1,000 groups.",
            ),
          groupBuckets.set(groupKey, {
            keys: keys,
            states: selectedMeasures.map(() => createSumAccumulator()),
          })),
          addAggregateValue(
            groupBuckets.get(groupKey).states[measureIndex],
            aggregateValue,
          ));
      }
    }
    return (
      (result.totals = totalAccumulators.map((state, measureIndex) =>
        evaluateAggregation(state, selectedMeasures[measureIndex]),
      )),
      (result.rows = modelQuery.groupBy.length
        ? [...groupBuckets.values()].map((group) => ({
            keys: group.keys,
            values: group.states.map((state, measureIndex) =>
              evaluateAggregation(state, selectedMeasures[measureIndex]),
            ),
          }))
        : [{ keys: [], values: [...result.totals] }]),
      (result.rows.length + 2) * result.headers.length >
        dataModelLimits.outputCells &&
        throwDaxModelError(
          "LIMIT",
          "The model result exceeds 20,000 output cells.",
        ),
      (result.complete = true),
      result
    );
  } catch (error) {
    return (
      (result.rows = []),
      (result.totals = []),
      result.diagnostics.push({
        code:
          error instanceof DaxModelError || error instanceof DaxError
            ? error.code
            : "MODEL",
        message:
          error instanceof Error ? error.message : "Model evaluation failed.",
        ...(error instanceof DaxModelError ? error.context : {}),
      }),
      result
    );
  }
}
function matchesComparisonOperator(cellValue, filter) {
  const value = filter.value;
  switch (filter.operator) {
    case "blank":
      return cellValue === null || cellValue === "";
    case "notBlank":
      return cellValue !== null && cellValue !== "";
    case "equals":
      return cellValue === value;
    case "notEquals":
      return cellValue !== value;
    case "gt":
      return cellValue !== null && cellValue > value;
    case "gte":
      return cellValue !== null && cellValue >= value;
    case "lt":
      return cellValue !== null && cellValue < value;
    case "lte":
      return cellValue !== null && cellValue <= value;
  }
}
function createSumAccumulator() {
  return {
    count: 0,
    sum: 0,
    correction: 0,
    min: 1 / 0,
    max: -1 / 0,
    distinct: new Set(),
  };
}
function addAggregateValue(state, value) {
  if (
    value !== null &&
    (state.count++,
    state.distinct.add(typedValueKey(value)),
    typeof value == "number")
  ) {
    const nextSum = state.sum + value;
    ((state.correction +=
      Math.abs(state.sum) >= Math.abs(value)
        ? state.sum - nextSum + value
        : value - nextSum + state.sum),
      (state.sum = nextSum),
      (state.min = Math.min(state.min, value)),
      (state.max = Math.max(state.max, value)));
  }
}
function evaluateAggregation(state, measure) {
  const compensatedSum = state.sum + state.correction;
  let result;
  switch (measure.operation) {
    case "count":
    case "countRows":
      result = state.count;
      break;
    case "distinctCount":
      result = state.distinct.size;
      break;
    case "sum":
      result = state.count ? compensatedSum : null;
      break;
    case "average":
      result = state.count ? compensatedSum / state.count : null;
      break;
    case "min":
      result = state.count ? state.min : null;
      break;
    case "max":
      result = state.count ? state.max : null;
      break;
    case "dax":
      return throwDaxModelError(
        "DAX_ENGINE",
        "DAX measures require the separate model expression evaluator.",
      );
  }
  return (
    result !== null &&
      !Number.isFinite(result) &&
      throwDaxModelError(
        "OVERFLOW",
        `Measure ${measure.name} exceeds finite numeric precision.`,
      ),
    result
  );
}
const dataModelResultCache = new WeakMap();
function previewDataModelOutput(workbook, request, options = {}) {
  const definition = validateDataModel(request.definition),
    query = validateModelQuery(request.query, definition),
    outputName = request.outputName.trim();
  if (
    !outputName ||
    outputName.length > 31 ||
    /[\\/?*\[\]:\u0000-\u001f]/.test(outputName)
  )
    throw Error(
      "Enter a valid output worksheet name of at most 31 characters.",
    );
  const preview = {
    request: { definition: definition, query: query, outputName: outputName },
    workbookHash: fingerprintCanonicalValue(workbook),
    contextHash: fingerprintCanonicalValue(
      normalizeModelHostPermissions(workbook, options),
    ),
    result: evaluateDataModelQuery(workbook, definition, query, options),
  };
  return (
    dataModelResultCache.set(preview, fingerprintCanonicalValue(preview)),
    preview
  );
}
function buildModelReportCells(worksheetRecord, outputRange) {
  const rangeAddress = parseRangeAddress(outputRange),
    cellRows = [];
  if (
    (rangeAddress.lastRow - rangeAddress.firstRow + 1) *
      (rangeAddress.lastCol - rangeAddress.firstCol + 1) >
    dataModelLimits.outputCells
  )
    throw Error("Model report range exceeds 20,000 cells.");
  for (
    let firstRow = rangeAddress.firstRow;
    firstRow <= rangeAddress.lastRow;
    firstRow++
  ) {
    const rowCells = [];
    for (
      let firstCol = rangeAddress.firstCol;
      firstCol <= rangeAddress.lastCol;
      firstCol++
    ) {
      const cell =
        worksheetRecord.cells[
          `${columnIndexToLetters(firstCol)}${firstRow + 1}`
        ];
      rowCells.push({
        raw: cell?.raw ?? null,
        literal: !!cell?.literal,
        note: cell?.note ?? "",
      });
    }
    cellRows.push(rowCells);
  }
  return fingerprintCanonicalValue(cellRows);
}
function publishDataModelReport(workbook, preview, mode, options = {}) {
  if (options.readOnly) throw Error("This workbook is read-only.");
  if (
    (assertWorkbookStructureEditable(workbook),
    !["save", "publish"].includes(mode))
  )
    throw Error("Choose Save model or Publish result.");
  if (
    dataModelResultCache.get(preview) !== fingerprintCanonicalValue(preview) ||
    preview.workbookHash !== fingerprintCanonicalValue(workbook) ||
    preview.contextHash !==
      fingerprintCanonicalValue(
        normalizeModelHostPermissions(workbook, options),
      )
  )
    throw Error(
      "The model, workbook or host permissions changed. Preview again.",
    );
  const queryResult = evaluateDataModelQuery(
    workbook,
    preview.request.definition,
    preview.request.query,
    options,
  );
  if (
    !queryResult.complete ||
    fingerprintCanonicalValue(queryResult) !==
      fingerprintCanonicalValue(preview.result)
  )
    throw Error(
      "Resolve model diagnostics and preview again. No partial model result is published.",
    );
  const workbookInput = normalizeWorkbookInput(workbook);
  if (
    ((workbookInput.dataModel = structuredClone(preview.request.definition)),
    mode === "publish")
  ) {
    if (!preview.request.query.measureIds.length)
      throw Error("Select at least one measure to publish.");
    if (
      workbookInput.sheets.some(
        (sheet) =>
          sheet.name.toLowerCase() === preview.request.outputName.toLowerCase(),
      )
    )
      throw Error("The output worksheet name is already in use.");
    inspectOutputNameSafety(workbook, preview.request.outputName);
    const worksheetRecord = createWorksheetRecord(preview.request.outputName),
      hasGroupBy = preview.request.query.groupBy.length > 0,
      outputRows = [
        queryResult.headers,
        ...queryResult.rows.map((row) => [...row.keys, ...row.values]),
      ];
    (hasGroupBy &&
      outputRows.push([
        "Grand total",
        ...Array(preview.request.query.groupBy.length - 1).fill(null),
        ...queryResult.totals,
      ]),
      outputRows.forEach((rowValues, rowIndex) =>
        rowValues.forEach((raw, columnIndex) => {
          worksheetRecord.cells[
            `${columnIndexToLetters(columnIndex)}${rowIndex + 1}`
          ] = {
            raw: raw,
            ...(typeof raw == "string" && raw.startsWith("=")
              ? { literal: true }
              : {}),
            ...(rowIndex === 0
              ? { style: { bold: true, fill: "#362C63", color: "#FFFFFF" } }
              : {}),
            ...(rowIndex > 0
              ? {
                  note: `Reviewed model snapshot: ${preview.request.definition.name}. Source hash ${queryResult.sourceHash}. Reopen Data Model to inspect source freshness and refresh.`,
                }
              : {}),
          };
        }),
      ),
      (worksheetRecord.columnWidths = Object.fromEntries(
        queryResult.headers.map((header, headerIndex) => [
          String(headerIndex),
          160,
        ]),
      )),
      (worksheetRecord.frozenRows = 1));
    const outputRange = `A1:${columnIndexToLetters(queryResult.headers.length - 1)}${outputRows.length}`;
    ((worksheetRecord.dataModelReport = {
      definition: structuredClone(preview.request.definition),
      query: structuredClone(preview.request.query),
      sourceHash: queryResult.sourceHash,
      definitionHash: fingerprintCanonicalValue(preview.request.definition),
      outputHash: buildModelReportCells(worksheetRecord, outputRange),
      outputRange: outputRange,
      createdAt: new Date().toISOString(),
      sourceRevision: workbook.revision,
      includeHidden: !!options.includeHidden,
      totals: [...queryResult.totals],
    }),
      workbookInput.sheets.push(worksheetRecord));
  }
  const normalizedWorkbook = normalizeWorkbookInput(workbookInput);
  return (dataModelResultCache.delete(preview), normalizedWorkbook);
}
function removeDataModelFromWorkbook(workbook, options = {}) {
  if (options.readOnly) throw Error("This workbook is read-only.");
  assertWorkbookStructureEditable(workbook);
  const workbookInput = normalizeWorkbookInput(workbook);
  return (
    delete workbookInput.dataModel,
    normalizeWorkbookInput(workbookInput)
  );
}
function getModelReportStatus(workbook, sheetId, options = {}) {
  const sheet = workbook.sheets.find((sheet) => sheet.id === sheetId),
    report = sheet?.dataModelReport;
  if (!sheet || !report) return { status: "none", message: "No model report." };
  try {
    if (
      !workbook.dataModel ||
      fingerprintCanonicalValue(workbook.dataModel) !== report.definitionHash
    )
      return {
        status: "stale",
        message:
          "Model definition changed. Preview and publish a fresh result.",
      };
    const queryResult = evaluateDataModelQuery(
      workbook,
      report.definition,
      report.query,
      {
        ...options,
        includeHidden: report.includeHidden && !!options.includeHidden,
      },
    );
    return queryResult.complete
      ? queryResult.sourceHash !== report.sourceHash ||
        buildModelReportCells(sheet, report.outputRange) !==
          report.outputHash ||
        fingerprintCanonicalValue(queryResult.totals) !==
          fingerprintCanonicalValue(report.totals)
        ? {
            status: "stale",
            message:
              "Model sources or output changed. Retained values are an earlier snapshot; preview and publish again.",
          }
        : {
            status: "current",
            message:
              "Model snapshot matches current sources and definition. Local fingerprints are not signed source evidence.",
          }
      : {
          status: "unavailable",
          message:
            "Model sources cannot be verified. Review relationships, source access and diagnostics.",
        };
  } catch {
    return {
      status: "unavailable",
      message:
        "Model source or output verification failed. Retained values are unverified.",
    };
  }
}
function DaxExpressionInput({ label, value, onChange }) {
  const { text } = useWorkbookMessages(),
    helpId = ne.useId(),
    textareaRef = ne.useRef(null);
  return i.jsxs("div", {
    className: "lg-dax-expression",
    children: [
      i.jsxs("label", {
        children: [
          i.jsx(LocalizedText, { children: label }),
          i.jsx("textarea", {
            ref: textareaRef,
            "aria-label": text(label),
            "aria-describedby": helpId,
            value,
            maxLength: daxExpressionLimits.expressionCharacters,
            rows: 3,
            spellCheck: false,
            onChange: (event) => onChange(event.target.value),
          }),
        ],
      }),
      i.jsx("div", {
        className: "lg-model-controls",
        children: i.jsxs("label", {
          children: [
            i.jsx(LocalizedText, { children: "Insert DAX function" }),
            i.jsxs("select", {
              "aria-label": text(`${label} insert function`),
              value: "",
              onChange: (event) => {
                if (!event.target.value) return;
                const selectionStart =
                    textareaRef.current?.selectionStart ?? value.length,
                  selectionEnd =
                    textareaRef.current?.selectionEnd ?? selectionStart,
                  insertText = event.target.value + "(";
                (onChange(
                  value.slice(0, selectionStart) +
                    insertText +
                    value.slice(selectionEnd),
                ),
                  requestAnimationFrame(() => {
                    (textareaRef.current?.focus(),
                      textareaRef.current?.setSelectionRange(
                        selectionStart + insertText.length,
                        selectionStart + insertText.length,
                      ));
                  }));
              },
              children: [
                i.jsx("option", {
                  value: "",
                  children: text("Choose a function"),
                }),
                Object.keys(daxFunctionNames).map((daxFunctionName) =>
                  i.jsx(
                    "option",
                    { value: daxFunctionName, children: daxFunctionName },
                    daxFunctionName,
                  ),
                ),
              ],
            }),
          ],
        }),
      }),
      i.jsxs("p", {
        id: helpId,
        className: "lg-model-help",
        children: [
          i.jsx(LocalizedText, {
            children:
              "Use table[column] and [Measure name]. This is DAX, evaluated separately from worksheet formulas.",
          }),
          i.jsx("br", {}),
          `SUM('Sales'[Amount]) · [Gross sales] - [Net sales] · CALCULATE([Net sales], Customers[Region] = "North")`,
        ],
      }),
    ],
  });
}
function DataModelDefinitionEditor({ definition, onChange }) {
  const { text } = useWorkbookMessages(),
    calculatedColumnCount = definition.tables.reduce(
      (total, table) =>
        total +
        table.columns.filter((column) => column.expression !== undefined)
          .length,
      0,
    ),
    updateTable = (tableId, changes) =>
      onChange({
        ...definition,
        tables: definition.tables.map((table) =>
          table.id === tableId ? { ...table, ...changes } : table,
        ),
      });
  return i.jsxs(i.Fragment, {
    children: [
      i.jsx("p", {
        className: "lg-model-help",
        children: i.jsx(LocalizedText, {
          children:
            "Calculated columns are evaluated for every source row before query filters. Sources remain unchanged. Preview is required before saving or publishing.",
        }),
      }),
      definition.tables.map((table) =>
        i.jsxs(
          "fieldset",
          {
            children: [
              i.jsx("legend", { children: table.name }),
              i.jsxs("label", {
                children: [
                  i.jsx(LocalizedText, { children: "Calendar date column" }),
                  i.jsxs("select", {
                    "aria-label": text(`${table.name} calendar date column`),
                    value: table.dateColumn ?? "",
                    onChange: (event) =>
                      updateTable(table.id, {
                        dateColumn: event.target.value || undefined,
                      }),
                    children: [
                      i.jsx("option", {
                        value: "",
                        children: text("Not a marked calendar"),
                      }),
                      table.columns
                        .filter(
                          (column) =>
                            column.expression === undefined &&
                            column.type === "date" &&
                            !column.nullable,
                        )
                        .map((column) =>
                          i.jsx(
                            "option",
                            { value: column.name, children: column.name },
                            column.name,
                          ),
                        ),
                    ],
                  }),
                ],
              }),
              i.jsx("p", {
                className: "lg-model-help",
                children: i.jsx(LocalizedText, {
                  children:
                    "TOTALYTD requires a marked table containing every day of one complete calendar year, without duplicate or blank dates.",
                }),
              }),
              table.columns.map((column, columnIndex) => {
                const calculatedColumnNumber = table.columns
                  .slice(0, columnIndex + 1)
                  .filter(
                    (candidateColumn) =>
                      candidateColumn.expression !== undefined,
                  ).length;
                return column.expression === undefined
                  ? null
                  : i.jsxs(
                      "fieldset",
                      {
                        children: [
                          i.jsxs("legend", {
                            children: [
                              i.jsx(LocalizedText, {
                                children: "Calculated column",
                              }),
                              " ",
                              calculatedColumnNumber,
                            ],
                          }),
                          i.jsxs("div", {
                            className: "lg-model-controls",
                            children: [
                              i.jsxs("label", {
                                children: [
                                  i.jsx(LocalizedText, { children: "Name" }),
                                  i.jsx("input", {
                                    "aria-label": text(
                                      `${table.name} calculated column ${calculatedColumnNumber} name`,
                                    ),
                                    value: column.name,
                                    onChange: (event) =>
                                      updateTable(table.id, {
                                        columns: table.columns.map(
                                          (column, position) =>
                                            position === columnIndex
                                              ? {
                                                  ...column,
                                                  name: event.target.value,
                                                }
                                              : column,
                                        ),
                                      }),
                                  }),
                                ],
                              }),
                              i.jsxs("label", {
                                children: [
                                  i.jsx(LocalizedText, { children: "Type" }),
                                  i.jsx("select", {
                                    "aria-label": text(
                                      `${table.name} calculated column ${calculatedColumnNumber} type`,
                                    ),
                                    value: column.type,
                                    onChange: (event) =>
                                      updateTable(table.id, {
                                        columns: table.columns.map(
                                          (column, position) =>
                                            position === columnIndex
                                              ? {
                                                  ...column,
                                                  type: event.target.value,
                                                }
                                              : column,
                                        ),
                                      }),
                                    children: [
                                      "number",
                                      "text",
                                      "boolean",
                                      "date",
                                    ].map((columnType) =>
                                      i.jsx(
                                        "option",
                                        {
                                          value: columnType,
                                          children: text(columnType),
                                        },
                                        columnType,
                                      ),
                                    ),
                                  }),
                                ],
                              }),
                              i.jsxs("label", {
                                children: [
                                  i.jsx("input", {
                                    type: "checkbox",
                                    "aria-label": text(
                                      `${table.name} calculated column ${calculatedColumnNumber} nullable`,
                                    ),
                                    checked: column.nullable,
                                    onChange: (event) =>
                                      updateTable(table.id, {
                                        columns: table.columns.map(
                                          (column, position) =>
                                            position === columnIndex
                                              ? {
                                                  ...column,
                                                  nullable:
                                                    event.target.checked,
                                                }
                                              : column,
                                        ),
                                      }),
                                  }),
                                  i.jsx(LocalizedText, {
                                    children: "Allow BLANK",
                                  }),
                                ],
                              }),
                              i.jsx("button", {
                                "aria-label": text(
                                  `Remove ${table.name} calculated column ${calculatedColumnNumber}`,
                                ),
                                onClick: () =>
                                  updateTable(table.id, {
                                    columns: table.columns.filter(
                                      (column, position) =>
                                        position !== columnIndex,
                                    ),
                                  }),
                                children: i.jsx(LocalizedText, {
                                  children: "Remove",
                                }),
                              }),
                            ],
                          }),
                          i.jsx(DaxExpressionInput, {
                            label: `${table.name} calculated column ${calculatedColumnNumber} expression`,
                            value: column.expression,
                            onChange: (expression) =>
                              updateTable(table.id, {
                                columns: table.columns.map(
                                  (column, position) =>
                                    position === columnIndex
                                      ? { ...column, expression: expression }
                                      : column,
                                ),
                              }),
                          }),
                        ],
                      },
                      columnIndex,
                    );
              }),
              i.jsx("button", {
                "aria-label": text(`Add calculated column to ${table.name}`),
                disabled:
                  calculatedColumnCount >=
                    daxExpressionLimits.calculatedColumns ||
                  table.columns.length >= dataModelLimits.columns,
                onClick: () => {
                  let nextNumber = 1;
                  for (
                    ;
                    table.columns.some(
                      (column) =>
                        column.name.toLowerCase() ===
                        `calculated ${nextNumber}`,
                    );
                  )
                    nextNumber++;
                  updateTable(table.id, {
                    columns: [
                      ...table.columns,
                      {
                        name: `Calculated ${nextNumber}`,
                        type: "number",
                        nullable: true,
                        expression: "BLANK()",
                      },
                    ],
                  });
                },
                children: i.jsx(LocalizedText, {
                  children: "Add calculated column",
                }),
              }),
            ],
          },
          table.id,
        ),
      ),
      !definition.tables.length &&
        i.jsx("p", {
          children: i.jsx(LocalizedText, {
            children: "Add a source in Tables first.",
          }),
        }),
    ],
  });
}
const LabeledField = (labelText, field) =>
  i.jsxs("label", {
    children: [i.jsx(LocalizedText, { children: labelText }), field],
  });

export {
  previewDataModelOutput,
  DataModelDefinitionEditor,
  publishDataModelReport,
  getModelReportStatus,
  removeDataModelFromWorkbook,
  LabeledField,
  buildModelTableFromRange,
  DaxExpressionInput,
};
