import {
  isPlainObjectValue,
  throwWithMessage,
} from "./01-cell-styles-and-themes.js";
import { isFiniteNumber } from "./04-pivot-tables-and-pictures.js";
import {
  string,
  number,
  boolean,
  null as _null,
  array,
  strictObject,
  discriminatedUnion,
  literal,
  union,
  enum as _enum,
  object,
  record,
} from "zod";
import {
  fingerprintCanonicalValue,
  sha256HexDigest,
} from "../01-schemas-date-functions-and-validation.js";
import { workbookLimits } from "../02-text-and-financial-functions.js";
import {
  columnLettersToIndex,
  sheetReferencePrefixPattern,
  unquotedNameCharPattern,
  offsetRowReference,
} from "./01-cell-styles-and-themes.js";

const analysisFormatVersion = "ledgerglass-analysis-v1",
  analysisLimits = Object.freeze({
    variables: 6,
    constraints: 16,
    candidates: 5e4,
    dependencyCells: 500,
    depth: 16,
    expressionCharacters: 4096,
    syntaxNodes: 512,
    sourceRows: 500,
    sourceColumns: 6,
    sourceCells: 2e3,
    outputCells: 2e4,
    metadataCharacters: 2e6,
  }),
  finiteNumberSchema = number().finite(),
  boundedNumberSchema = finiteNumberSchema.min(-1e9).max(1e9),
  cellReferenceSchema = strictObject({
    sheetId: string().min(1).max(100),
    address: string().regex(/^[A-Z]{1,3}[1-9]\d{0,6}$/),
  }),
  outputWorksheetNameSchema = string()
    .trim()
    .min(1)
    .max(31)
    .refine(
      (worksheetName) => !/[\\/?*\[\]:\u0000-\u001f]/.test(worksheetName),
      "Invalid output worksheet name",
    ),
  solverAnalysisSpecSchema = strictObject({
    kind: literal("solver"),
    outputName: outputWorksheetNameSchema,
    objective: cellReferenceSchema,
    direction: _enum(["maximize", "minimize", "target"]),
    target: boundedNumberSchema.optional(),
    variables: array(
      strictObject({
        cell: cellReferenceSchema,
        lower: boundedNumberSchema,
        upper: boundedNumberSchema,
        domain: _enum(["continuous", "integer", "binary"]),
      }),
    )
      .min(1)
      .max(6),
    constraints: array(
      strictObject({
        left: cellReferenceSchema,
        operator: _enum(["<=", ">=", "="]),
        right: union([boundedNumberSchema, cellReferenceSchema]),
      }),
    ).max(16),
    tolerance: finiteNumberSchema.min(1e-10).max(1e-5).default(1e-8),
    maxCandidates: number().int().min(1).max(5e4).default(5e4),
  }),
  statisticsAnalysisSpecSchema = strictObject({
    kind: literal("statistics"),
    outputName: outputWorksheetNameSchema,
    sheetId: string().min(1).max(100),
    range: string().min(1).max(50),
    method: _enum([
      "descriptive",
      "correlation",
      "covariance",
      "regression",
      "histogram",
      "distribution",
      "forecast",
    ]),
    missing: _enum(["reject", "exclude"]).default("reject"),
    bins: array(boundedNumberSchema).min(1).max(50).optional(),
    distribution: _enum(["normal", "binomial", "poisson"]).optional(),
    mean: boundedNumberSchema.optional(),
    standardDeviation: finiteNumberSchema.positive().max(1e9).optional(),
    trials: number().int().min(0).max(1e3).optional(),
    probability: finiteNumberSchema.min(0).max(1).optional(),
    lambda: finiteNumberSchema.positive().max(1e3).optional(),
    forecast: _enum(["linear", "moving-average", "exponential"]).optional(),
    horizon: number().int().min(1).max(24).optional(),
    window: number().int().min(2).max(100).optional(),
    alpha: finiteNumberSchema.positive().max(1).optional(),
  }),
  analysisSpecSchema = discriminatedUnion("kind", [
    solverAnalysisSpecSchema,
    statisticsAnalysisSpecSchema,
  ]),
  analysisCellValueSchema = union([
    string().max(32767),
    finiteNumberSchema,
    boolean(),
    _null(),
  ]),
  analysisOutcomeSchema = strictObject({
    status: _enum([
      "complete",
      "optimal",
      "infeasible",
      "work-limit",
      "precision-limit",
    ]),
    success: boolean(),
    message: string().max(2e3),
    algorithm: string().max(300),
    assumptions: array(string().max(1e3)).max(30),
    exclusions: array(
      strictObject({
        row: number().int().positive(),
        reason: string().max(200),
        locator: string().max(300),
      }),
    ).max(500),
    included: number().int().nonnegative(),
    inspected: number().int().nonnegative(),
    iterations: number().int().nonnegative(),
    rows: array(array(analysisCellValueSchema).max(16)).max(3e3),
    solution: array(finiteNumberSchema).max(6).optional(),
    objective: finiteNumberSchema.optional(),
    baseline: array(finiteNumberSchema).max(6).optional(),
  }),
  sk = string().regex(/^[a-f0-9]{64}$/),
  analysisReportRecordSchema = strictObject({
    engine: literal(analysisFormatVersion),
    request: analysisSpecSchema,
    result: analysisOutcomeSchema,
    definitionHash: sk,
    sourceHash: sk,
    outputHash: sk,
    outputRange: string().max(50),
    createdAt: string().datetime(),
    sourceRevision: number().int().nonnegative(),
    includeHidden: boolean(),
    applied: boolean(),
  });
function parseAnalysisRequest(request) {
  const parseResult = analysisSpecSchema.safeParse(request);
  if (!parseResult.success)
    throw Error(
      `Invalid analysis request: ${parseResult.error.issues[0]?.path.join(".")}: ${parseResult.error.issues[0]?.message}`,
    );
  const data = parseResult.data;
  if (data.kind === "solver") {
    if (data.direction === "target" && data.target === undefined)
      throw Error("Enter a finite target value.");
    if (
      new Set(data.variables.map((variable) => JSON.stringify(variable.cell)))
        .size !== data.variables.length
    )
      throw Error("Changing cells must be unique.");
    if (
      data.variables.some(
        (variable) =>
          variable.lower > variable.upper ||
          (variable.domain === "binary" &&
            (variable.lower < 0 || variable.upper > 1)),
      )
    )
      throw Error(
        "Variable bounds must be ordered; binary bounds must lie within zero and one.",
      );
    if (
      data.variables.some((variable) => variable.domain === "continuous") &&
      data.variables.some((variable) => variable.domain !== "continuous")
    )
      throw Error(
        "Mixed integer and continuous models are outside this selected Solver subset.",
      );
  }
  return data;
}
function validateAnalysisReport(storedWorkbook, workbook) {
  const t = storedWorkbook;
  let metadataLength = 0;
  for (const sheet of workbook.sheets) {
    const analysisReport = t.sheets.find(
      (storedSheet) => storedSheet.id === sheet.id,
    )?.analysisReport;
    if (analysisReport === undefined) continue;
    if (
      ((metadataLength += JSON.stringify(analysisReport).length),
      metadataLength > analysisLimits.metadataCharacters)
    )
      throw Error("Analysis metadata exceeds 2,000,000 characters.");
    const reportParseResult =
      analysisReportRecordSchema.safeParse(analysisReport);
    if (!reportParseResult.success)
      throw Error(
        `Invalid analysis report: ${reportParseResult.error.issues[0]?.message}`,
      );
    const data = reportParseResult.data;
    if (
      ((data.request = parseAnalysisRequest(data.request)),
      data.definitionHash !== fingerprintCanonicalValue(data.request) ||
        !data.result.success ||
        !["complete", "optimal"].includes(data.result.status) ||
        data.result.rows.reduce((total, row) => total + row.length, 0) >
          analysisLimits.outputCells)
    )
      throw Error("Analysis report metadata is inconsistent.");
    const rangeMatch =
        /^([A-Z]{1,3})([1-9]\d{0,6}):([A-Z]{1,3})([1-9]\d{0,6})$/.exec(
          data.outputRange,
        ),
      columnLettersToNumber = (letters) =>
        [...letters].reduce(
          (columnNumber, letter) =>
            columnNumber * 26 + letter.charCodeAt(0) - 64,
          0,
        );
    if (
      !rangeMatch ||
      columnLettersToNumber(rangeMatch[3]) > 16384 ||
      Number(rangeMatch[4]) > 1048576 ||
      columnLettersToNumber(rangeMatch[3]) <
        columnLettersToNumber(rangeMatch[1]) ||
      Number(rangeMatch[4]) < Number(rangeMatch[2]) ||
      (columnLettersToNumber(rangeMatch[3]) -
        columnLettersToNumber(rangeMatch[1]) +
        1) *
        (Number(rangeMatch[4]) - Number(rangeMatch[2]) + 1) >
        analysisLimits.outputCells
    )
      throw Error("Invalid or oversized analysis output range.");
    if (
      (data.request.kind === "solver" &&
        (data.result.status !== "optimal" ||
          data.result.solution?.length !== data.request.variables.length ||
          data.result.baseline?.length !== data.request.variables.length ||
          data.result.objective === undefined)) ||
      (data.request.kind === "statistics" &&
        (data.result.status !== "complete" || data.applied))
    )
      throw Error("Analysis result does not match its request.");
    sheet.analysisReport = data;
  }
}
const reportOutputFormatVersion = "ledgerglass-report-output-v1",
  reportOutputLimits = Object.freeze({
    outputs: 20,
    cells: 2e3,
    rows: 500,
    columns: 20,
    metadataBytes: 16e6,
    imagePixels: 16e6,
    imageWidth: 4096,
    imageHeight: 16384,
    inventories: 20,
    inventoryItems: 1e3,
  }),
  createBoundedTextSchema = (maxLength) => string().max(maxLength),
  nonEmptyLabelSchema = string().min(1).max(200),
  reportHashSchema = string().regex(/^[a-f0-9]{64}$/),
  AU = string().datetime();
const hashRecordExcludingHash = (record) => {
  const { hash, ...recordWithoutHash } = record;
  return fingerprintCanonicalValue(recordWithoutHash);
};
const daxExpressionLimits = Object.freeze({
    expressionCharacters: 4096,
    syntaxNodes: 512,
    depth: 48,
    calculatedColumns: 16,
  }),
  daxFunctionNames = Object.freeze({
    BLANK: [0, 0],
    TRUE: [0, 0],
    FALSE: [0, 0],
    ISBLANK: [1, 1],
    IF: [2, 3],
    COALESCE: [2, 16],
    DIVIDE: [2, 3],
    SUM: [1, 1],
    AVERAGE: [1, 1],
    MIN: [1, 1],
    MAX: [1, 1],
    DISTINCTCOUNT: [1, 1],
    COUNTROWS: [1, 1],
    SUMX: [2, 2],
    CALCULATE: [1, 17],
    FILTER: [2, 2],
    ALL: [1, 1],
    RELATED: [1, 1],
    SELECTEDVALUE: [1, 2],
    TOTALYTD: [2, 2],
  });
class DaxError extends Error {
  constructor(code, message) {
    (super(message), (this.code = code), (this.name = "DaxError"));
  }
}
const throwDaxError = (code, message) => {
  throw new DaxError(code, `DAX: ${message}`);
};
function parseDaxExpression(expression) {
  if (
    typeof expression != "string" ||
    !expression.trim() ||
    expression.length > daxExpressionLimits.expressionCharacters
  )
    return throwDaxError(
      "DAX_LIMIT",
      "Enter an expression of 1–4,096 characters.",
    );
  const source = expression.trim().replace(/^=/, ""),
    tokens = [];
  let position = 0;
  const readDelimited = (label, delimiter, kind) => {
    const startPosition = position++;
    let tokenText = "",
      isClosed = false;
    for (; position < source.length;) {
      const character = source[position++];
      if (character === delimiter)
        if (source[position] === delimiter)
          ((tokenText += delimiter), position++);
        else {
          isClosed = true;
          break;
        }
      else tokenText += character;
    }
    (isClosed ||
      throwDaxError(
        "DAX_SYNTAX",
        `Unclosed ${label} at position ${startPosition + 1}.`,
      ),
      tokens.push({ kind: kind, value: tokenText, position: startPosition }));
  };
  for (; position < source.length;) {
    if (/\s/.test(source[position])) {
      position++;
      continue;
    }
    if (source[position] === '"') {
      readDelimited('"', '"', "string");
      continue;
    }
    if (source[position] === "'") {
      readDelimited("'", "'", "quoted");
      continue;
    }
    if (source[position] === "[") {
      readDelimited("[", "]", "bracket");
      continue;
    }
    const remaining = source.slice(position),
      numberMatch = /^(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?/.exec(remaining),
      identifierMatch = /^[A-Za-z_][A-Za-z0-9_]*/.exec(remaining),
      operatorMatch = /^(?:==|<>|<=|>=|&&|\|\||[+\-*/^&=<>(),])/.exec(
        remaining,
      ),
      tokenMatch = numberMatch ?? identifierMatch ?? operatorMatch;
    (tokenMatch ||
      throwDaxError(
        "DAX_UNSUPPORTED",
        `Unsupported syntax at position ${position + 1}. Use the documented DAX subset; worksheet and query syntax are not evaluated.`,
      ),
      tokens.push({
        kind: numberMatch ? "number" : identifierMatch ? "identifier" : "op",
        value: tokenMatch[0],
        position: position,
      }),
      (position += tokenMatch[0].length),
      tokens.length > daxExpressionLimits.syntaxNodes * 3 &&
        throwDaxError("DAX_LIMIT", "Expression token budget exceeded."));
  }
  tokens.push({ kind: "end", value: "", position: position });
  let tokenIndex = 0,
    nodeCount = 0;
  const peekToken = () => tokens[tokenIndex],
    nextToken = () => tokens[tokenIndex++],
    expectToken = (expected) => {
      nextToken().value !== expected &&
        throwDaxError("DAX_SYNTAX", `Expected ${expected}.`);
    },
    countNode = (parsedNode) => (
      ++nodeCount > daxExpressionLimits.syntaxNodes &&
        throwDaxError("DAX_LIMIT", "Expression exceeds 512 syntax nodes."),
      parsedNode
    ),
    binaryPrecedence = {
      "||": 1,
      "&&": 2,
      "=": 4,
      "==": 4,
      "<>": 4,
      "<": 4,
      ">": 4,
      "<=": 4,
      ">=": 4,
      "&": 5,
      "+": 6,
      "-": 6,
      "*": 7,
      "/": 7,
      "^": 9,
    },
    parseExpressionNode = (minPrecedence = 0, depth = 0) => {
      depth > daxExpressionLimits.depth &&
        throwDaxError("DAX_LIMIT", "Expression nesting exceeds 48 levels.");
      const token = nextToken();
      let left;
      if (token.kind === "number") {
        const numericValue = Number(token.value);
        (Number.isFinite(numericValue) ||
          throwDaxError(
            "DAX_NUMBER",
            "Numeric literal exceeds finite precision.",
          ),
          (left = countNode({ kind: "literal", value: numericValue })));
      } else if (token.kind === "string")
        left = countNode({ kind: "literal", value: token.value });
      else if (token.value === "(")
        ((left = parseExpressionNode(0, depth + 1)), expectToken(")"));
      else if (
        token.value === "+" ||
        token.value === "-" ||
        (token.kind === "identifier" && token.value.toUpperCase() === "NOT")
      )
        left = countNode({
          kind: "unary",
          op: token.value.toUpperCase(),
          value: parseExpressionNode(
            token.value.toUpperCase() === "NOT" ? 3 : 8,
            depth + 1,
          ),
        });
      else if (token.kind === "bracket")
        left = countNode({ kind: "measure", name: token.value });
      else if (token.kind === "identifier" || token.kind === "quoted")
        if (peekToken().value === "(" && token.kind === "identifier") {
          const functionName = token.value.toUpperCase();
          (Object.prototype.hasOwnProperty.call(
            daxFunctionNames,
            functionName,
          ) ||
            throwDaxError(
              "DAX_UNSUPPORTED",
              `Unsupported function ${token.value}. No worksheet fallback is used.`,
            ),
            nextToken());
          const args = [];
          if (peekToken().value !== ")")
            do {
              if (
                (args.push(parseExpressionNode(0, depth + 1)),
                peekToken().value !== ",")
              )
                break;
              nextToken();
            } while (true);
          expectToken(")");
          const argumentRange = daxFunctionNames[functionName];
          ((args.length < argumentRange[0] || args.length > argumentRange[1]) &&
            throwDaxError(
              "DAX_ARGUMENT",
              `${functionName} expects ${argumentRange[0] === argumentRange[1] ? argumentRange[0] : `${argumentRange[0]}–${argumentRange[1]}`} arguments in this subset.`,
            ),
            (left = countNode({
              kind: "call",
              name: functionName,
              args: args,
            })));
        } else
          peekToken().kind === "bracket"
            ? (left = countNode({
                kind: "field",
                table: token.value,
                column: nextToken().value,
              }))
            : token.kind === "identifier" &&
                ["TRUE", "FALSE"].includes(token.value.toUpperCase())
              ? (left = countNode({
                  kind: "literal",
                  value: token.value.toUpperCase() === "TRUE",
                }))
              : (left = countNode({ kind: "table", name: token.value }));
      else
        return throwDaxError(
          "DAX_SYNTAX",
          `Expected an expression at position ${token.position + 1}.`,
        );
      for (
        ;
        peekToken().kind === "op" &&
        binaryPrecedence[peekToken().value] !== undefined &&
        binaryPrecedence[peekToken().value] >= minPrecedence;
      ) {
        const value = nextToken().value,
          precedence = binaryPrecedence[value];
        left = countNode({
          kind: "binary",
          op: value,
          left: left,
          right: parseExpressionNode(precedence + 1, depth + 1),
        });
      }
      return left;
    },
    rootNode = parseExpressionNode();
  return (
    peekToken().kind !== "end" &&
      throwDaxError(
        "DAX_UNSUPPORTED",
        `Unexpected ${peekToken().value}. VAR, RETURN, EVALUATE, worksheet references and unsupported syntax are not evaluated.`,
      ),
    rootNode
  );
}
function walkDaxSyntaxTree(syntaxNode, visitor) {
  (visitor(syntaxNode),
    syntaxNode.kind === "call"
      ? syntaxNode.args.forEach((arg) => walkDaxSyntaxTree(arg, visitor))
      : syntaxNode.kind === "binary"
        ? (walkDaxSyntaxTree(syntaxNode.left, visitor),
          walkDaxSyntaxTree(syntaxNode.right, visitor))
        : syntaxNode.kind === "unary" &&
          walkDaxSyntaxTree(syntaxNode.value, visitor));
}
const columnarModelFormatVersion = "ledgerglass-columnar-v1",
  dataModelLimits = Object.freeze({
    tables: 8,
    columns: 50,
    rowsPerTable: 500,
    cellsPerTable: 2e3,
    modelCells: 16e3,
    relationships: 12,
    measures: 16,
    filters: 16,
    groupFields: 3,
    groups: 1e3,
    outputCells: 2e4,
    operations: 2e6,
    metadataCharacters: 2e6,
  }),
  tableIdentifierSchema = string()
    .min(1)
    .max(100)
    .regex(/^[A-Za-z0-9_.-]+$/),
  modelNameSchema = string()
    .trim()
    .min(1)
    .max(128)
    .refine((modelName) => !/[\u0000-\u001f\u007f]/.test(modelName)),
  modelColumnSchema = strictObject({
    name: modelNameSchema,
    type: _enum(["text", "number", "boolean", "date"]),
    nullable: boolean(),
    expression: string()
      .trim()
      .min(1)
      .max(daxExpressionLimits.expressionCharacters)
      .optional(),
  }),
  modelTableSourceSchema = strictObject({
    id: tableIdentifierSchema,
    name: modelNameSchema,
    sheetId: string().min(1).max(100),
    range: string().min(1).max(50),
    tableId: string().min(1).max(100).optional(),
    columns: array(modelColumnSchema).min(1).max(dataModelLimits.columns),
    dateColumn: modelNameSchema.optional(),
  }),
  modelGroupFieldSchema = strictObject({
    tableId: tableIdentifierSchema,
    column: modelNameSchema,
  }),
  relationshipOptionsSchema = strictObject({
    id: tableIdentifierSchema,
    one: modelGroupFieldSchema,
    many: modelGroupFieldSchema,
    cardinality: _enum(["one-to-many", "one-to-one"]),
    direction: literal("single"),
  }),
  modelMeasureSchema = strictObject({
    id: tableIdentifierSchema,
    name: modelNameSchema,
    tableId: tableIdentifierSchema,
    column: modelNameSchema.optional(),
    operation: _enum([
      "sum",
      "count",
      "countRows",
      "average",
      "min",
      "max",
      "distinctCount",
      "dax",
    ]),
    expression: string()
      .trim()
      .min(1)
      .max(daxExpressionLimits.expressionCharacters)
      .optional(),
  }),
  dataModelDefinitionSchema = strictObject({
    engine: literal(columnarModelFormatVersion),
    name: modelNameSchema,
    tables: array(modelTableSourceSchema).min(1).max(dataModelLimits.tables),
    relationships: array(relationshipOptionsSchema).max(
      dataModelLimits.relationships,
    ),
    measures: array(modelMeasureSchema).max(dataModelLimits.measures),
  }),
  modelCellValueSchema = union([
    string().max(32767),
    number().finite(),
    boolean(),
    _null(),
  ]),
  modelFilterSchema = strictObject({
    ...modelGroupFieldSchema.shape,
    operator: _enum([
      "equals",
      "notEquals",
      "gt",
      "gte",
      "lt",
      "lte",
      "blank",
      "notBlank",
    ]),
    value: modelCellValueSchema.optional(),
  }),
  modelQuerySchema = strictObject({
    groupBy: array(modelGroupFieldSchema).max(dataModelLimits.groupFields),
    filters: array(modelFilterSchema).max(dataModelLimits.filters),
    measureIds: array(tableIdentifierSchema).max(dataModelLimits.measures),
  }),
  assertUniqueCaseInsensitive = (values, label) => {
    if (
      new Set(values.map((value) => value.toLowerCase())).size !== values.length
    )
      throw Error(`${label} must be unique, including case variants.`);
  };
function validateDataModel(rawDefinition) {
  const parseResult = dataModelDefinitionSchema.safeParse(rawDefinition);
  if (!parseResult.success)
    throw Error(
      `Invalid data model: ${parseResult.error.issues[0]?.path.join(".")}: ${parseResult.error.issues[0]?.message}`,
    );
  const data = parseResult.data;
  if (
    (assertUniqueCaseInsensitive(
      data.tables.map((table) => table.id),
      "Model table IDs",
    ),
    assertUniqueCaseInsensitive(
      data.tables.map((table) => table.name),
      "Model table names",
    ),
    assertUniqueCaseInsensitive(
      data.relationships.map((relationship) => relationship.id),
      "Relationship IDs",
    ),
    assertUniqueCaseInsensitive(
      data.measures.map((measure) => measure.id),
      "Measure IDs",
    ),
    assertUniqueCaseInsensitive(
      data.measures.map((measure) => measure.name),
      "Measure names",
    ),
    data.tables.reduce(
      (total, tableDefinition) =>
        total +
        tableDefinition.columns.filter(
          (column) => column.expression !== undefined,
        ).length,
      0,
    ) > daxExpressionLimits.calculatedColumns)
  )
    throw Error("The model supports at most 16 DAX calculated columns.");
  for (const table of data.tables) {
    if (
      (assertUniqueCaseInsensitive(
        table.columns.map((column) => column.name),
        `Columns in ${table.name}`,
      ),
      !table.columns.some((column) => column.expression === undefined))
    )
      throw Error(
        "A model table requires at least one physical source column.",
      );
    if (
      table.dateColumn &&
      !table.columns.some(
        (column) =>
          column.name === table.dateColumn &&
          column.type === "date" &&
          !column.nullable &&
          column.expression === undefined,
      )
    )
      throw Error(
        "A marked calendar requires a nonnullable physical date column.",
      );
  }
  const resolveField = (fieldReference) => {
    const table = data.tables.find(
        (table) => table.id === fieldReference.tableId,
      ),
      matchingColumn = table?.columns.find(
        (column) => column.name === fieldReference.column,
      );
    if (!matchingColumn)
      throw Error(
        `Unknown model field ${fieldReference.tableId}.${fieldReference.column}.`,
      );
    return matchingColumn;
  };
  for (const relationship of data.relationships) {
    if (relationship.one.tableId === relationship.many.tableId)
      throw Error(
        "Self relationships are unsupported; use distinct role-playing tables.",
      );
    if (
      resolveField(relationship.one).type !==
      resolveField(relationship.many).type
    )
      throw Error(
        `Relationship ${relationship.id} requires matching key types.`,
      );
    if (
      resolveField(relationship.one).expression !== undefined ||
      resolveField(relationship.many).expression !== undefined
    )
      throw Error(
        "Relationship keys must be physical source columns, not DAX calculated columns.",
      );
  }
  for (const measure of data.measures) {
    if (!data.tables.some((table) => table.id === measure.tableId))
      throw Error(`Unknown measure table ${measure.tableId}.`);
    if (measure.operation === "dax") {
      if (!measure.expression || measure.column !== undefined)
        throw Error(
          "A DAX measure requires an expression and no aggregate column.",
        );
    } else {
      if (measure.expression !== undefined)
        throw Error(
          "Unrecognized expression for an aggregate measure; only DAX measures may contain it.",
        );
      if (measure.operation === "countRows") {
        if (measure.column !== undefined)
          throw Error("Count rows does not take a column.");
      } else {
        const measureColumn = resolveField({
          tableId: measure.tableId,
          column: measure.column ?? "",
        });
        if (
          ["sum", "average", "min", "max"].includes(measure.operation) &&
          measureColumn.type !== "number"
        )
          throw Error(`Measure ${measure.name} requires a number column.`);
      }
    }
  }
  return data;
}
function validateModelQuery(query, definition) {
  const parseResult = modelQuerySchema.safeParse(query);
  if (!parseResult.success)
    throw Error(`Invalid model query: ${parseResult.error.issues[0]?.message}`);
  const data = parseResult.data;
  (assertUniqueCaseInsensitive(data.measureIds, "Selected measures"),
    assertUniqueCaseInsensitive(
      data.groupBy.map((groupField) => JSON.stringify(groupField)),
      "Group fields",
    ));
  for (const field of [...data.groupBy, ...data.filters])
    if (
      !definition.tables
        .find((table) => table.id === field.tableId)
        ?.columns.some((column) => column.name === field.column)
    )
      throw Error(`Unknown query field ${field.tableId}.${field.column}.`);
  if (
    data.measureIds.some(
      (measureId) =>
        !definition.measures.some((measure) => measure.id === measureId),
    )
  )
    throw Error("A selected model measure is missing.");
  return data;
}
const lk = string().regex(/^[a-f0-9]{64}$/),
  dataModelReportSchema = strictObject({
    definition: dataModelDefinitionSchema,
    query: modelQuerySchema,
    sourceHash: lk,
    definitionHash: lk,
    outputHash: lk,
    outputRange: string().min(1).max(50),
    createdAt: string().datetime(),
    sourceRevision: number().int().nonnegative(),
    includeHidden: boolean(),
    totals: array(number().finite().nullable()).max(dataModelLimits.measures),
  });
function validateDataModelReport(storedWorkbook, workbook) {
  const t = storedWorkbook;
  let metadataLength = 0;
  if (t.dataModel !== undefined) {
    if (
      ((metadataLength += JSON.stringify(t.dataModel).length),
      metadataLength > dataModelLimits.metadataCharacters)
    )
      throw Error("Data model metadata exceeds 2,000,000 characters.");
    workbook.dataModel = validateDataModel(t.dataModel);
  }
  for (const sheet of workbook.sheets) {
    const dataModelReport = t.sheets.find(
      (storedSheet) => storedSheet.id === sheet.id,
    )?.dataModelReport;
    if (dataModelReport === undefined) continue;
    if (
      ((metadataLength += JSON.stringify(dataModelReport).length),
      metadataLength > dataModelLimits.metadataCharacters)
    )
      throw Error("Data model metadata exceeds 2,000,000 characters.");
    const parseResult = dataModelReportSchema.safeParse(dataModelReport);
    if (!parseResult.success)
      throw Error(
        `Invalid model report: ${parseResult.error.issues[0]?.message}`,
      );
    const data = parseResult.data;
    if (
      ((data.definition = validateDataModel(data.definition)),
      (data.query = validateModelQuery(data.query, data.definition)),
      data.definitionHash !== fingerprintCanonicalValue(data.definition) ||
        data.totals.length !== data.query.measureIds.length)
    )
      throw Error(
        "Model report definition or measure metadata is inconsistent.",
      );
    sheet.dataModelReport = data;
  }
}
const sourceImportLimits = Object.freeze({
    bytes: 4 * 1024 * 1024,
    retainedBytes: 12 * 1024 * 1024,
    imports: 10,
    tables: 20,
    rows: 1e3,
    columns: 100,
    cells: 2e4,
    depth: 32,
    nodes: 1e5,
    confidence: 0.9,
  }),
  documentSourceFormats = ["json", "xml", "pdf", "png", "jpeg", "html"],
  retainedSourceHashSchema = string().regex(/^[a-f0-9]{64}$/),
  IU = union([
    createBoundedTextSchema(32767),
    number().finite(),
    boolean(),
    _null(),
  ]),
  sourceRegionSchema = object({
    page: number().int().min(1).max(1e4),
    x: number().min(0).max(1),
    y: number().min(0).max(1),
    width: number().positive().max(1),
    height: number().positive().max(1),
  })
    .strict()
    .refine(
      (region) =>
        region.x + region.width <= 1.000001 &&
        region.y + region.height <= 1.000001,
      "Source region must fit its page.",
    );
function Poe(jsonText, expectedKind) {
  if (typeof jsonText != "string") return false;
  try {
    const parsedJson = JSON.parse(jsonText);
    return expectedKind === "array"
      ? Array.isArray(parsedJson)
      : !!parsedJson &&
          typeof parsedJson == "object" &&
          !Array.isArray(parsedJson);
  } catch {
    return false;
  }
}
const nativeCellValueSchema = object({
    value: IU,
    type: _enum(["string", "number", "boolean", "null", "array", "object"]),
    path: createBoundedTextSchema(2048),
    lexical: createBoundedTextSchema(32767).optional(),
    declaredType: createBoundedTextSchema(200).optional(),
    missing: boolean().optional(),
    confidence: number().min(0).max(1),
    region: sourceRegionSchema.optional(),
  })
    .strict()
    .refine(
      (cell) =>
        cell.type === "array" || cell.type === "object"
          ? Poe(cell.value, cell.type)
          : cell.type === "null"
            ? cell.value === null
            : typeof cell.value === cell.type,
      "Native value and type disagree.",
    )
    .refine(
      (cell) => !cell.missing || (cell.value === null && cell.type === "null"),
      "Missing fields must be null placeholders.",
    ),
  extractedTableSchema = object({
    id: createBoundedTextSchema(200).min(1),
    name: createBoundedTextSchema(200).min(1),
    path: createBoundedTextSchema(2048),
    ambiguous: boolean(),
    columns: array(
      object({
        name: createBoundedTextSchema(200).min(1),
        path: createBoundedTextSchema(2048),
      }).strict(),
    )
      .min(1)
      .max(sourceImportLimits.columns),
    rows: array(array(nativeCellValueSchema).max(sourceImportLimits.columns))
      .min(1)
      .max(sourceImportLimits.rows),
  })
    .strict()
    .superRefine((table, refinementContext) => {
      (new Set(table.columns.map((column) => column.path)).size !==
        table.columns.length ||
        table.rows.some((row) => row.length !== table.columns.length)) &&
        refinementContext.addIssue({
          code: "custom",
          message: "Table columns and rows must align with unique paths.",
        });
    }),
  retainedSourceFileSchema = object({
    fileName: createBoundedTextSchema(255)
      .min(1)
      .refine(
        (fileName) =>
          !/[\\/\x00-\x1f\x7f\u202a-\u202e\u2066-\u2069]/.test(fileName),
      ),
    format: _enum(documentSourceFormats),
    byteLength: number().int().min(1).max(sourceImportLimits.bytes),
    sha256: retainedSourceHashSchema,
    base64: createBoundedTextSchema(
      Math.ceil(sourceImportLimits.bytes / 3) * 4,
    ),
  }).strict(),
  extractedDocumentSchema = object({
    schemaVersion: literal(1),
    fileName: retainedSourceFileSchema.shape.fileName,
    format: _enum(documentSourceFormats),
    byteLength: retainedSourceFileSchema.shape.byteLength,
    sha256: retainedSourceHashSchema,
    tables: array(extractedTableSchema).max(sourceImportLimits.tables),
    method: _enum(["native-json", "native-xml", "document-adapter"]),
    extractor: createBoundedTextSchema(200).min(1),
    pageCount: number().int().min(1).max(1e4).optional(),
    warnings: array(createBoundedTextSchema(2048)).max(30),
  }).strict(),
  reviewCorrectionSchema = object({
    decision: _enum(["accept", "exclude", "correct"]),
    value: IU.optional(),
  })
    .strict()
    .refine(
      (correction) =>
        correction.decision !== "correct" || correction.value !== undefined,
      "A corrected value is required.",
    ),
  sourceImportRecordSchema = object({
    id: createBoundedTextSchema(100).min(1),
    sheetId: createBoundedTextSchema(100).min(1),
    sheetName: createBoundedTextSchema(31).min(1),
    importedAt: string().datetime(),
    original: retainedSourceFileSchema,
    method: extractedDocumentSchema.shape.method,
    extractor: createBoundedTextSchema(200).min(1),
    pageCount: number().int().min(1).max(1e4).optional(),
    table: extractedTableSchema,
    resolutions: record(string().regex(/^\d+:\d+$/), reviewCorrectionSchema),
    ambiguityReviewed: boolean(),
  }).strict();
function validateExtractedSourceTables(rawDocument) {
  const extractedDocument = extractedDocumentSchema.parse(rawDocument);
  let cellCount = 0;
  if (
    new Set(extractedDocument.tables.map((table) => table.id)).size !==
    extractedDocument.tables.length
  )
    throw Error("Source table IDs must be unique.");
  for (const table of extractedDocument.tables)
    for (const row of table.rows)
      for (const cell of row) {
        if (++cellCount > sourceImportLimits.cells)
          throw Error("Source cell budget exceeded.");
        if (
          extractedDocument.method === "document-adapter" &&
          (!cell.region ||
            !extractedDocument.pageCount ||
            cell.region.page > extractedDocument.pageCount ||
            (["png", "jpeg"].includes(extractedDocument.format) &&
              extractedDocument.pageCount !== 1))
        )
          throw Error(
            "Document cells require a source page and bounding region.",
          );
      }
  if (
    (extractedDocument.format === "json" &&
      extractedDocument.method !== "native-json") ||
    (extractedDocument.format === "xml" &&
      extractedDocument.method !== "native-xml") ||
    (!["json", "xml"].includes(extractedDocument.format) &&
      extractedDocument.method !== "document-adapter")
  )
    throw Error("Source format and extraction method disagree.");
  return extractedDocument;
}
function validateRetainedSourceFile(original) {
  const retainedFile = retainedSourceFileSchema.parse(original),
    extension = retainedFile.fileName.split(".").at(-1)?.toLowerCase();
  if (
    (extension === "jpg"
      ? "jpeg"
      : extension === "htm"
        ? "html"
        : extension) !== retainedFile.format
  )
    throw Error("Source extension and format disagree.");
  if (
    retainedFile.base64.length !== Math.ceil(retainedFile.byteLength / 3) * 4 ||
    !/^[A-Za-z0-9+/]*={0,2}$/.test(retainedFile.base64)
  )
    throw Error("Invalid retained source encoding.");
  const binaryString = atob(retainedFile.base64);
  if (
    btoa(binaryString) !== retainedFile.base64 ||
    binaryString.length !== retainedFile.byteLength
  )
    throw Error("Retained source byte length changed.");
  const bytes = Uint8Array.from(binaryString, (character) =>
    character.charCodeAt(0),
  );
  if (sha256HexDigest(bytes) !== retainedFile.sha256)
    throw Error("Retained source checksum mismatch.");
  return bytes;
}
function buildRetainedSourceRecord(bytes, fileName, format) {
  let binaryString = "";
  for (let offset = 0; offset < bytes.length; offset += 8192)
    binaryString += String.fromCharCode(
      ...bytes.subarray(offset, offset + 8192),
    );
  const original = {
    fileName: fileName,
    format: format,
    byteLength: bytes.length,
    sha256: sha256HexDigest(bytes),
    base64: btoa(binaryString),
  };
  return (validateRetainedSourceFile(original), original);
}
function resolveExtractedCellValue(extractedCell, correction) {
  if (extractedCell.confidence < sourceImportLimits.confidence && !correction)
    throw Error("Review every low-confidence cell before importing.");
  if (!correction) return extractedCell.value;
  const parsedCorrection = reviewCorrectionSchema.parse(correction);
  return parsedCorrection.decision === "exclude"
    ? null
    : parsedCorrection.decision === "correct"
      ? parsedCorrection.value
      : extractedCell.value;
}
const uk = new WeakSet();
function MU(value) {
  Object.freeze(value);
  for (const nestedValue of Object.values(value))
    nestedValue && typeof nestedValue == "object" && MU(nestedValue);
}
function validateSourceImports(workbook, target) {
  if (workbook.sourceImports === undefined) return;
  if (
    !Array.isArray(workbook.sourceImports) ||
    workbook.sourceImports.length > sourceImportLimits.imports
  )
    throw Error("Source import count exceeded.");
  if (
    workbook.sourceImports.some(
      (sourceImport) =>
        !sourceImport ||
        typeof sourceImport != "object" ||
        !uk.has(sourceImport),
    ) &&
    new TextEncoder().encode(JSON.stringify(workbook.sourceImports)).length >
      24 * 1024 * 1024
  )
    throw Error("Source provenance metadata budget exceeded.");
  let retainedBytes = 0;
  const seenImportIds = new Set();
  if (
    ((target.sourceImports = workbook.sourceImports.map((sourceImport) => {
      if (
        sourceImport &&
        typeof sourceImport == "object" &&
        uk.has(sourceImport)
      ) {
        const retainedImport = sourceImport;
        if (
          ((retainedBytes += retainedImport.original.byteLength),
          seenImportIds.has(retainedImport.id))
        )
          throw Error("Duplicate source import ID.");
        return (seenImportIds.add(retainedImport.id), retainedImport);
      }
      const sourceImportRecord = sourceImportRecordSchema.parse(sourceImport);
      if (
        (validateRetainedSourceFile(sourceImportRecord.original),
        validateExtractedSourceTables({
          schemaVersion: 1,
          fileName: sourceImportRecord.original.fileName,
          format: sourceImportRecord.original.format,
          byteLength: sourceImportRecord.original.byteLength,
          sha256: sourceImportRecord.original.sha256,
          tables: [sourceImportRecord.table],
          method: sourceImportRecord.method,
          extractor: sourceImportRecord.extractor,
          ...(sourceImportRecord.pageCount
            ? { pageCount: sourceImportRecord.pageCount }
            : {}),
          warnings: [],
        }),
        sourceImportRecord.table.ambiguous &&
          !sourceImportRecord.ambiguityReviewed)
      )
        throw Error("Ambiguous source tables require review.");
      const cellKeys = new Set();
      if (
        (sourceImportRecord.table.rows.forEach((row, rowIndex) =>
          row.forEach((cell, columnIndex) => {
            const cellKey = `${rowIndex}:${columnIndex}`;
            (cellKeys.add(cellKey),
              resolveExtractedCellValue(
                cell,
                sourceImportRecord.resolutions[cellKey],
              ));
          }),
        ),
        Object.keys(sourceImportRecord.resolutions).some(
          (resolution) => !cellKeys.has(resolution),
        ))
      )
        throw Error("Unknown source review cell.");
      if (seenImportIds.has(sourceImportRecord.id))
        throw Error("Duplicate source import ID.");
      return (
        seenImportIds.add(sourceImportRecord.id),
        (retainedBytes += sourceImportRecord.original.byteLength),
        MU(sourceImportRecord),
        uk.add(sourceImportRecord),
        sourceImportRecord
      );
    })),
    retainedBytes > sourceImportLimits.retainedBytes)
  )
    throw Error("Retained source byte budget exceeded.");
}
const xlsxInspectionLimits = Object.freeze({
    maxBytes: 10 * 1024 * 1024,
    maxParts: 2e3,
    maxDetails: 1e5,
  }),
  q5 = new WeakSet();
function validateRetainedXlsxSource(xlsxSource) {
  if (
    !xlsxSource ||
    typeof xlsxSource != "object" ||
    xlsxSource.schemaVersion !== 1 ||
    typeof xlsxSource.fileName != "string" ||
    !/\.xlsx$/i.test(xlsxSource.fileName) ||
    xlsxSource.fileName.length > 255 ||
    /[\\/\x00-\x1f\x7f]/.test(xlsxSource.fileName) ||
    !Number.isSafeInteger(xlsxSource.byteLength) ||
    xlsxSource.byteLength < 22 ||
    xlsxSource.byteLength > xlsxInspectionLimits.maxBytes ||
    typeof xlsxSource.sha256 != "string" ||
    !/^[a-f0-9]{64}$/.test(xlsxSource.sha256) ||
    typeof xlsxSource.base64 != "string" ||
    xlsxSource.base64.length !== Math.ceil(xlsxSource.byteLength / 3) * 4 ||
    !/^[A-Za-z0-9+/]*={0,2}$/.test(xlsxSource.base64)
  )
    throw Error(
      "Invalid retained XLSX source metadata or 10 MiB source limit exceeded.",
    );
  let binaryString;
  try {
    binaryString = atob(xlsxSource.base64);
  } catch {
    throw Error("Invalid retained XLSX source encoding.");
  }
  if (
    btoa(binaryString) !== xlsxSource.base64 ||
    binaryString.length !== xlsxSource.byteLength
  )
    throw Error("Retained XLSX source byte length or encoding changed.");
  const bytes = Uint8Array.from(binaryString, (character) =>
    character.charCodeAt(0),
  );
  if (sha256HexDigest(bytes) !== xlsxSource.sha256)
    throw Error("Retained XLSX source checksum does not match its bytes.");
  return bytes;
}
function retainOriginalXlsx(bytes, fileName) {
  if (bytes.byteLength > xlsxInspectionLimits.maxBytes)
    throw Error("Original XLSX exceeds the 10 MiB retention limit.");
  let binaryString = "";
  for (let offset = 0; offset < bytes.length; offset += 8192)
    binaryString += String.fromCharCode(
      ...bytes.subarray(offset, offset + 8192),
    );
  const xlsxSource = {
    schemaVersion: 1,
    fileName: fileName,
    byteLength: bytes.length,
    sha256: sha256HexDigest(bytes),
    base64: btoa(binaryString),
  };
  return (
    validateRetainedXlsxSource(xlsxSource),
    Object.freeze(xlsxSource),
    q5.add(xlsxSource),
    xlsxSource
  );
}
function readRetainedXlsxSource(workbook, target) {
  if (workbook.xlsxSource === undefined) return;
  const xlsxSource = workbook.xlsxSource;
  if (xlsxSource && typeof xlsxSource == "object" && q5.has(xlsxSource)) {
    target.xlsxSource = xlsxSource;
    return;
  }
  (validateRetainedXlsxSource(xlsxSource),
    (target.xlsxSource = Object.freeze({
      schemaVersion: 1,
      fileName: xlsxSource.fileName,
      byteLength: xlsxSource.byteLength,
      sha256: xlsxSource.sha256,
      base64: xlsxSource.base64,
    })),
    q5.add(target.xlsxSource));
}
function requireRetainedXlsxSource(workbook) {
  if (!workbook.xlsxSource)
    throw Error(
      "This workbook has no retained XLSX source. Open an XLSX file to retain its original bytes.",
    );
  return {
    bytes: validateRetainedXlsxSource(workbook.xlsxSource),
    fileName: workbook.xlsxSource.fileName,
    sha256: workbook.xlsxSource.sha256,
  };
}
const alternateFormatExtensions = [
    "xls",
    "xlsb",
    "xlsm",
    "xltx",
    "xltm",
    "ods",
  ],
  isSafeFileName = (fileName) =>
    typeof fileName == "string" &&
    fileName.length > 0 &&
    fileName.length <= 255 &&
    !/[\\/\x00-\x1f\x7f\u202a-\u202e\u2066-\u2069]/.test(fileName),
  jU = new WeakSet();
function validateAlternateSource(alternateSource) {
  if (
    !alternateSource ||
    typeof alternateSource != "object" ||
    !alternateFormatExtensions.includes(alternateSource.format) ||
    !isSafeFileName(alternateSource.fileName) ||
    !alternateSource.fileName
      .toLowerCase()
      .endsWith("." + alternateSource.format) ||
    !alternateSource.conversion ||
    !["template-adapter", "external-conversion"].includes(
      alternateSource.conversion.method,
    ) ||
    (alternateSource.conversion.method === "template-adapter" &&
      alternateSource.format !== "xltx") ||
    !isSafeFileName(alternateSource.conversion.convertedFileName) ||
    !/\.xlsx$/i.test(alternateSource.conversion.convertedFileName) ||
    typeof alternateSource.conversion.convertedSha256 != "string" ||
    !/^[a-f0-9]{64}$/.test(alternateSource.conversion.convertedSha256) ||
    typeof alternateSource.conversion.reviewedAt != "string" ||
    !/^\d{4}-\d\d-\d\dT\d\d:\d\d:\d\d\.\d{3}Z$/.test(
      alternateSource.conversion.reviewedAt,
    ) ||
    !Number.isFinite(Date.parse(alternateSource.conversion.reviewedAt)) ||
    alternateSource.conversion.equivalence !== "not-verified"
  )
    throw Error("Invalid alternate workbook source or conversion provenance.");
  return validateRetainedXlsxSource({
    ...alternateSource,
    fileName: "retained.xlsx",
  });
}
function retainAlternateFormatSource(bytes, fileName, format, conversion) {
  const alternateSource = {
    ...retainOriginalXlsx(bytes, "retained.xlsx"),
    fileName: fileName,
    format: format,
    conversion: { ...conversion },
  };
  return (
    validateAlternateSource(alternateSource),
    Object.freeze(alternateSource.conversion),
    Object.freeze(alternateSource),
    jU.add(alternateSource),
    alternateSource
  );
}
function normalizeAlternateFormatSource(workbook, target) {
  if (workbook.alternateSource === undefined) return;
  const alternateSource = workbook.alternateSource;
  if (jU.has(alternateSource)) {
    target.alternateSource = alternateSource;
    return;
  }
  const validatedSource = validateAlternateSource(alternateSource);
  target.alternateSource = retainAlternateFormatSource(
    validatedSource,
    alternateSource.fileName,
    alternateSource.format,
    alternateSource.conversion,
  );
}
function requireAlternateFormatOriginal(workbook) {
  const alternateSource = workbook.alternateSource;
  if (!alternateSource)
    throw Error("This workbook has no retained alternate-format original.");
  return {
    bytes: validateAlternateSource(alternateSource),
    fileName: alternateSource.fileName,
    sha256: alternateSource.sha256,
    format: alternateSource.format,
  };
}
const automationLimits = Object.freeze({
    maxFileBytes: 10 * 1024 * 1024,
    maxSourceBytes: 256 * 1024,
    maxArtifactBytes: 2 * 1024 * 1024,
    maxStoredBytes: 10 * 1024 * 1024,
    maxArtifacts: 20,
    maxRuns: 100,
    maxCommands: 100,
    maxCells: 500,
    maxTokens: 2e4,
  }),
  automationArtifactKinds = [
    "vba",
    "office-script",
    "office-script-container",
    "excel-container",
    "legacy-container",
  ];
function encodeBytesToBase64(bytes) {
  let binaryString = "";
  for (let offset = 0; offset < bytes.length; offset += 8192)
    binaryString += String.fromCharCode(
      ...bytes.subarray(offset, offset + 8192),
    );
  return btoa(binaryString);
}
function decodeAutomationArtifact(base64) {
  if (
    typeof base64 != "string" ||
    base64.length > Math.ceil(automationLimits.maxArtifactBytes / 3) * 4 ||
    !/^(?:[A-Za-z0-9+/]{4})*(?:[A-Za-z0-9+/]{2}==|[A-Za-z0-9+/]{3}=)?$/.test(
      base64,
    )
  )
    throw Error("Invalid automation artifact encoding or size.");
  const binaryString = atob(base64);
  if (btoa(binaryString) !== base64)
    throw Error("Automation artifact is not canonical base64.");
  return Uint8Array.from(binaryString, (character) => character.charCodeAt(0));
}
const requireAutomationText = (value, maxLength, label) => {
    if (typeof value != "string" || !value.trim() || value.length > maxLength)
      throw Error(`Invalid automation ${label}.`);
    return value;
  },
  requireAutomationSourceHash = (value) => {
    if (typeof value != "string" || !/^[a-f0-9]{64}$/.test(value))
      throw Error("Invalid automation source hash.");
    return value;
  },
  requireAutomationEvidenceCount = (value, minimum, maximum) => {
    if (
      !Number.isSafeInteger(value) ||
      Number(value) < minimum ||
      Number(value) > maximum
    )
      throw Error("Invalid automation evidence count.");
    return Number(value);
  },
  requireAutomationTimestamp = (value) => {
    const timestampText = requireAutomationText(value, 40, "timestamp");
    if (
      !/^\d{4}-\d\d-\d\dT/.test(timestampText) ||
      !Number.isFinite(Date.parse(timestampText))
    )
      throw Error("Invalid automation timestamp.");
    return timestampText;
  },
  validateAutomationApprovedRange = (value) => {
    const rangeText = requireAutomationText(value, 40, "approved range"),
      rangeMatch = /^([A-Z]{1,3})([1-9]\d*):([A-Z]{1,3})([1-9]\d*)$/.exec(
        rangeText,
      );
    if (!rangeMatch) throw Error("Invalid automation approved range.");
    const columnNumber = (letters) =>
        [...letters].reduce(
          (total, letter) => total * 26 + letter.charCodeAt(0) - 64,
          0,
        ),
      [startColumn, startRow, endColumn, endRow] = [
        columnNumber(rangeMatch[1]),
        Number(rangeMatch[2]),
        columnNumber(rangeMatch[3]),
        Number(rangeMatch[4]),
      ];
    if (
      startColumn > endColumn ||
      startRow > endRow ||
      endColumn > 16384 ||
      endRow > 1048576 ||
      (endColumn - startColumn + 1) * (endRow - startRow + 1) >
        automationLimits.maxCells
    )
      throw Error("Invalid automation approved range.");
    return rangeText;
  };
function requireAutomationFileName(value) {
  const fileName = requireAutomationText(value, 255, "file name");
  if (
    /[\\/\x00-\x1f\x7f]/.test(fileName) ||
    fileName === "." ||
    fileName === ".."
  )
    throw Error("Invalid automation file name.");
  return fileName;
}
function validateAutomationArtifacts(workbook, target) {
  if (workbook.automationArtifacts !== undefined) {
    if (
      !Array.isArray(workbook.automationArtifacts) ||
      workbook.automationArtifacts.length > automationLimits.maxArtifacts
    )
      throw Error("Keep at most 20 automation artifacts.");
    let retainedBytes = 0;
    const seenArtifactIds = new Set();
    target.automationArtifacts = workbook.automationArtifacts.map(
      (automationArtifact) => {
        if (
          !automationArtifact ||
          typeof automationArtifact != "object" ||
          Array.isArray(automationArtifact)
        )
          throw Error("Invalid automation artifact.");
        const artifact = automationArtifact,
          artifactId = requireAutomationText(artifact.id, 100, "artifact ID");
        if (seenArtifactIds.has(artifactId))
          throw Error("Duplicate automation artifact ID.");
        seenArtifactIds.add(artifactId);
        const bytes = decodeAutomationArtifact(artifact.base64);
        if (
          ((retainedBytes += bytes.length),
          !bytes.length ||
            bytes.length !== artifact.byteLength ||
            bytes.length > automationLimits.maxArtifactBytes ||
            retainedBytes > automationLimits.maxStoredBytes)
        )
          throw Error(
            "Automation artifacts exceed their 2 MB per-file or 10 MB retained-byte budget.",
          );
        if (
          sha256HexDigest(bytes) !==
          requireAutomationSourceHash(artifact.sha256)
        )
          throw Error(
            "Automation artifact checksum does not match its preserved bytes.",
          );
        if (!automationArtifactKinds.includes(artifact.kind))
          throw Error("Unknown automation artifact kind.");
        return {
          id: artifactId,
          fileName: requireAutomationFileName(artifact.fileName),
          kind: artifact.kind,
          sha256: artifact.sha256,
          byteLength: bytes.length,
          base64: artifact.base64,
          retainedAt: requireAutomationTimestamp(artifact.retainedAt),
          retainedBy: requireAutomationText(
            artifact.retainedBy,
            100,
            "retainer",
          ),
        };
      },
    );
  }
  if (workbook.automationRuns !== undefined) {
    if (
      !Array.isArray(workbook.automationRuns) ||
      workbook.automationRuns.length > automationLimits.maxRuns
    )
      throw Error("Automation run evidence is limited to 100 runs.");
    const seenRunIds = new Set();
    target.automationRuns = workbook.automationRuns.map((automationRun) => {
      if (
        !automationRun ||
        typeof automationRun != "object" ||
        Array.isArray(automationRun)
      )
        throw Error("Invalid automation run evidence.");
      const run = automationRun,
        runId = requireAutomationText(run.id, 100, "run ID");
      if (seenRunIds.has(runId)) throw Error("Duplicate automation run ID.");
      if (
        (seenRunIds.add(runId),
        !["vba", "office-script"].includes(run.language) ||
          run.adapter !== "literal-ranges-v1")
      )
        throw Error("Unknown automation migration adapter.");
      if (
        !Array.isArray(run.sourceLines) ||
        run.sourceLines.length > automationLimits.maxCommands
      )
        throw Error("Invalid automation source lines.");
      const beforeRevision = requireAutomationEvidenceCount(
          run.beforeRevision,
          0,
          Number.MAX_SAFE_INTEGER - 2,
        ),
        afterRevision = requireAutomationEvidenceCount(
          run.afterRevision,
          1,
          Number.MAX_SAFE_INTEGER - 1,
        );
      if (afterRevision !== beforeRevision + 1)
        throw Error("Automation run revisions are inconsistent.");
      if (run.sourceLines.length !== run.commands)
        throw Error(
          "Automation source-line evidence must match the command count.",
        );
      return {
        id: runId,
        eventId: requireAutomationText(run.eventId, 100, "event ID"),
        sourceName: requireAutomationFileName(run.sourceName),
        sourceSha256: requireAutomationSourceHash(run.sourceSha256),
        codeSha256: requireAutomationSourceHash(run.codeSha256),
        language: run.language,
        adapter: run.adapter,
        actor: requireAutomationText(run.actor, 100, "actor"),
        timestamp: requireAutomationTimestamp(run.timestamp),
        beforeRevision: beforeRevision,
        afterRevision: afterRevision,
        sheetId: requireAutomationText(run.sheetId, 100, "sheet ID"),
        range: validateAutomationApprovedRange(run.range),
        commands: requireAutomationEvidenceCount(
          run.commands,
          1,
          automationLimits.maxCommands,
        ),
        changedCells: requireAutomationEvidenceCount(
          run.changedCells,
          0,
          automationLimits.maxCells,
        ),
        sourceLines: run.sourceLines.map((sourceLine) =>
          requireAutomationEvidenceCount(
            sourceLine,
            1,
            automationLimits.maxSourceBytes + 1,
          ),
        ),
      };
    });
  }
}
function validateDelimitedSourceFootprint(delimitedSource, target) {
  if (delimitedSource === undefined) return;
  if (
    !delimitedSource ||
    typeof delimitedSource != "object" ||
    Array.isArray(delimitedSource)
  )
    throw Error("Invalid delimited source footprint.");
  const footprint = delimitedSource;
  if (
    !Number.isInteger(footprint.records) ||
    Number(footprint.records) < 1 ||
    Number(footprint.records) > workbookLimits.maxRows ||
    !Number.isInteger(footprint.columns) ||
    Number(footprint.columns) < 1 ||
    Number(footprint.columns) > workbookLimits.maxCols ||
    typeof footprint.hasHeader != "boolean"
  )
    throw Error(
      "Delimited source footprint exceeds worksheet bounds or has invalid header metadata.",
    );
  target.delimitedSource = {
    records: Number(footprint.records),
    columns: Number(footprint.columns),
    hasHeader: footprint.hasHeader,
  };
}
const B2 = 100,
  referenceImpactKinds = [
    "deleted-cells",
    "deleted-sheet",
    "overwritten",
    "missing-sheet",
    "out-of-bounds",
  ],
  HO = (object, allowedKeys) =>
    Object.keys(object).every((key) => allowedKeys.includes(key)),
  isBoundedPlainText = (text, maxLength) =>
    typeof text == "string" &&
    !!text.trim() &&
    text.length <= maxLength &&
    !/[\u0000-\u001f]/.test(text);
const whatIfLimits = {
    inputs: 32,
    outputs: 16,
    scenarios: 100,
    axisValues: 100,
    trials: 400,
    dependencyCells: 500,
    workUnits: 5e4,
    savedAnalyses: 40,
  },
  cellAddressKey = (cellAddress) =>
    JSON.stringify([cellAddress.sheetId, cellAddress.address]);
function validateWhatIfFields(fields, allowedFields) {
  if (
    !fields ||
    typeof fields != "object" ||
    Array.isArray(fields) ||
    ![Object.prototype, null].includes(Object.getPrototypeOf(fields)) ||
    Object.keys(fields).some((field) => !allowedFields.includes(field))
  )
    throw Error("Invalid what-if analysis fields.");
  return fields;
}
function requireWhatIfName(value, maxLength = 100) {
  if (
    typeof value != "string" ||
    !value.trim() ||
    value.length > maxLength ||
    /[\u0000-\u001f]/.test(value)
  )
    throw Error(
      `What-if names must be nonempty text of at most ${maxLength} characters.`,
    );
  return value.trim();
}
function requireWhatIfList(value, maxEntries, minEntries = 1) {
  if (
    !Array.isArray(value) ||
    value.length < minEntries ||
    value.length > maxEntries
  )
    throw Error(`What-if lists require ${minEntries}–${maxEntries} entries.`);
  return Array.from(value);
}
function validateWhatIfValues(values, maxEntries) {
  return requireWhatIfList(values, maxEntries).map((value) => {
    if (!isFiniteNumber(value))
      throw Error("Every what-if input value must be a finite number.");
    return value;
  });
}
function expandWhatIfScenarios(request) {
  return request.kind === "scenarios"
    ? request.scenarios.map((scenario) => ({
        name: scenario.name,
        inputs: scenario.values,
      }))
    : request.kind === "one-variable"
      ? request.values.map((value) => ({
          name: String(value),
          inputs: [value],
        }))
      : request.columnValues.flatMap((columnValue) =>
          request.rowValues.map((rowValue) => ({
            name: `${rowValue} / ${columnValue}`,
            inputs: [rowValue, columnValue],
          })),
        );
}
const printRenderLimits = {
    areaCells: 5e3,
    pages: 100,
    titleRows: 100,
    titleColumns: 50,
    breaksPerAxis: 100,
    renderedCells: 1e5,
  },
  defaultPrintHeaderFooter = {
    header: `&[Workbook]
Engagement: &[Engagement] · Period: &[Period]`,
    footer: "Revision &[Revision] · &A · Page &P of &N",
  },
  headerFooterTokenPattern =
    /&&|&\[(Workbook|Revision|Engagement|Period)\]|&[PNA]/g;
function resolvePrintIdentityToken(token, workbook) {
  return token === "Workbook"
    ? (workbook?.name ?? "[Workbook not available]")
    : token === "Revision"
      ? workbook
        ? String(workbook.revision)
        : "[Revision not available]"
      : token === "Engagement"
        ? workbook?.metadata?.engagement || "[Engagement not set]"
        : workbook?.metadata?.period || "[Period not set]";
}
function expandHeaderFooterTemplate(
  template,
  workbook,
  sheet,
  pageNumber,
  total,
) {
  return template.replace(headerFooterTokenPattern, (match, identityToken) =>
    identityToken
      ? resolvePrintIdentityToken(identityToken, workbook)
      : match === "&&"
        ? "&"
        : match === "&A"
          ? sheet.name
          : String(match === "&P" ? pageNumber : total),
  );
}
function exportHeaderFooterTemplate(template, workbookInput) {
  const exported = template.replace(
    headerFooterTokenPattern,
    (match, identityToken) => {
      if (!identityToken) return match;
      if (!workbookInput)
        throw Error(
          "Exporting print identity fields requires the current workbook snapshot.",
        );
      return resolvePrintIdentityToken(identityToken, workbookInput).replace(
        /&/g,
        "&&",
      );
    },
  );
  if (exported.length > 255)
    throw Error(
      "Resolved print identifiers exceed Excel's 255-character header/footer limit. Shorten the header, footer or workbook properties before exporting.",
    );
  return exported;
}
function validatePrintPageBreaks(breaks, axisLength) {
  if (breaks !== undefined) {
    if (
      !Array.isArray(breaks) ||
      breaks.length > printRenderLimits.breaksPerAxis ||
      breaks.some(
        (breakValue, index) =>
          !Number.isSafeInteger(breakValue) ||
          breakValue < 1 ||
          breakValue >= axisLength ||
          (index > 0 && breakValue <= breaks[index - 1]),
      )
    )
      throw Error(
        "Manual page breaks require at most 100 distinct ascending boundaries within each worksheet axis (before row/column 2 or later).",
      );
    return breaks.length ? [...breaks] : undefined;
  }
}

function verifyPrintSnapshotGuard(workbook, printSetup) {
  if (
    printSetup.workbookId !== workbook.id ||
    printSetup.revision !== workbook.revision ||
    printSetup.snapshotGuard !== JSON.stringify(workbook)
  )
    throw Error(
      "The workbook changed after page setup opened. Close and reopen page setup.",
    );
}
function resolvePageDimensions(pageSetup) {
  const portraitSize =
    pageSetup.paper === "a4"
      ? [793.7007874015749, 1122.5196850393702]
      : pageSetup.paper === "legal"
        ? [816, 1344]
        : [816, 1056];
  return pageSetup.orientation === "landscape"
    ? { width: portraitSize[1], height: portraitSize[0] }
    : { width: portraitSize[0], height: portraitSize[1] };
}
const textCleanupFunctionNames = [
    "TRIM",
    "CLEAN",
    "UPPER",
    "LOWER",
    "PROPER",
    "LEFT",
    "RIGHT",
    "MID",
    "LEN",
    "FIND",
    "SEARCH",
    "SUBSTITUTE",
    "REPLACE",
    "VALUE",
    "NUMBERVALUE",
    "TEXT",
    "CONCAT",
    "TEXTJOIN",
    "EXACT",
    "TEXTBEFORE",
    "TEXTAFTER",
  ],
  throwTextCleanupEvidenceError = () => {
    throw Error(
      "Invalid or over-capacity text cleanup evidence. Keep the original native backup.",
    );
  };
function assertPlainObjectKeys(value, allowedKeys) {
  (!isPlainObjectValue(value) ||
    Object.keys(value).some((key) => !allowedKeys.includes(key))) &&
    throwTextCleanupEvidenceError();
}
function WU(arg) {
  return (
    arg === null ||
    typeof arg == "boolean" ||
    (typeof arg == "number" && Number.isFinite(arg)) ||
    (typeof arg == "string" && arg.length <= 32767)
  );
}
function jm(text, maxLength) {
  return typeof text == "string" && text.length > 0 && text.length <= maxLength;
}
function validateTextCleanupSteps(steps) {
  if (!Array.isArray(steps) || !steps.length || steps.length > 12)
    throw Error("Choose 1–12 text cleanup steps.");
  return steps.map(
    (step) => (
      assertPlainObjectKeys(step, ["function", "args"]),
      (!textCleanupFunctionNames.includes(step.function) ||
        !Array.isArray(step.args) ||
        step.args.length > 5 ||
        step.args.some((arg) => arg === null || !WU(arg)) ||
        JSON.stringify(step.args).length > 4096) &&
        throwTextCleanupEvidenceError(),
      { function: step.function, args: [...step.args] }
    ),
  );
}
function assertTextCleanupEvidenceUnchanged(workbook, nextState) {
  if (!workbook.textCleanups?.length) return nextState;
  const savedCleanupsById = new Map(
    workbook.textCleanups.map((textCleanup) => [textCleanup.id, textCleanup]),
  );
  for (const nextCleanup of nextState.textCleanups ?? []) {
    const savedCleanup = savedCleanupsById.get(nextCleanup.id);
    if (
      savedCleanup &&
      JSON.stringify(savedCleanup) !== JSON.stringify(nextCleanup)
    )
      throw Error("Saved original text cleanup evidence cannot be changed.");
  }
  return {
    ...nextState,
    textCleanups: [
      ...workbook.textCleanups,
      ...(nextState.textCleanups ?? []).filter(
        (cleanup) => !savedCleanupsById.has(cleanup.id),
      ),
    ],
  };
}
const textInferenceLimits = {
    inferencePatterns: 8e3,
    inferenceCandidates: 30,
    evidenceCharacters: 5e6,
  },
  columnIndexSchema = number().int().min(0).max(49),
  distinctColumnPositions = array(columnIndexSchema)
    .min(1)
    .max(50)
    .refine(
      (positions) => new Set(positions).size === positions.length,
      "Choose distinct column positions.",
    ),
  cellTextValueSchema = string().max(32767),
  transformCellValueSchema = union([
    cellTextValueSchema,
    number().finite(),
    boolean(),
    _null(),
  ]),
  textCaseModeSchema = _enum(["none", "upper", "lower", "proper"]),
  textExtractStepSchema = strictObject({
    column: columnIndexSchema,
    mode: _enum(["whole", "token", "slice", "before", "after"]),
    delimiter: string().min(1).max(16).default(" "),
    index: number().int().min(-50).max(49).default(0),
    start: number().int().min(0).max(32767).default(0),
    length: number().int().min(1).max(32767).default(1),
  }),
  textJoinSpecSchema = strictObject({
    parts: array(textExtractStepSchema).min(1).max(8),
    separator: string().max(100),
    prefix: string().max(1e3),
    suffix: string().max(1e3),
    casing: textCaseModeSchema,
  }),
  sourceRowValueSchema = strictObject({
    row: number().int().min(1).max(1048576),
    value: cellTextValueSchema,
  }),
  distinctExampleRowsSchema = array(sourceRowValueSchema)
    .min(2)
    .max(10)
    .refine(
      (rows) =>
        new Set(rows.map((sourceRow) => sourceRow.row)).size === rows.length,
      "Use examples from distinct source rows.",
    ),
  fieldTransformStepSchema = discriminatedUnion("kind", [
    strictObject({
      kind: literal("trim"),
      columns: distinctColumnPositions,
      mode: _enum(["edges", "collapse"]),
    }),
    strictObject({ kind: literal("clean"), columns: distinctColumnPositions }),
    strictObject({
      kind: literal("case"),
      columns: distinctColumnPositions,
      mode: textCaseModeSchema,
    }),
    strictObject({
      kind: literal("replace"),
      columns: distinctColumnPositions,
      find: string().min(1).max(1e3),
      replacement: string().max(1e3),
    }),
    strictObject({
      kind: literal("split"),
      column: columnIndexSchema,
      mode: _enum(["delimited", "fixed"]),
      delimiters: string().max(16),
      collapse: boolean(),
      breaks: array(number().int().min(1).max(32766)).max(49),
    }),
    strictObject({
      kind: literal("concat"),
      columns: distinctColumnPositions,
      separator: string().max(100),
      name: string().min(1).max(100),
      skipBlank: boolean(),
    }),
    strictObject({
      kind: literal("convert"),
      columns: distinctColumnPositions,
      to: _enum(["text", "number", "integer", "boolean", "date"]),
      decimal: _enum([".", ","]),
      group: _enum(["", ".", ",", " ", "'"]),
      dateOrder: _enum(["MDY", "DMY"]),
    }),
    strictObject({
      kind: literal("fill"),
      columns: distinctColumnPositions,
      direction: _enum(["down", "up", "value"]),
      value: transformCellValueSchema,
      groupColumns: array(columnIndexSchema).max(50),
    }),
    strictObject({ kind: literal("select"), columns: distinctColumnPositions }),
    strictObject({
      kind: literal("pattern"),
      name: string().min(1).max(100),
      pattern: textJoinSpecSchema,
      examples: distinctExampleRowsSchema,
      inferenceColumns: distinctColumnPositions.max(8),
    }),
  ]),
  fieldTransformRecipeSchema = strictObject({
    sourceSheetId: string().min(1).max(100),
    sourceRange: string().min(1).max(50),
    header: boolean(),
    outputName: string().min(1).max(31),
    steps: array(fieldTransformStepSchema).max(12),
  });
function validateFieldTransformRecipe(request) {
  const parseResult = fieldTransformRecipeSchema.safeParse(request);
  if (!parseResult.success)
    throw Error(
      `Invalid field transformation: ${parseResult.error.issues[0]?.message ?? "review the recipe"}`,
    );
  const data = parseResult.data;
  if (
    ((data.outputName = data.outputName.trim()),
    !data.outputName ||
      /[\u0000-\u001f\u007f\\/\[\]*?:]/.test(data.outputName) ||
      /^'|'$/.test(data.outputName))
  )
    throw Error("Choose a valid output worksheet name of 1–31 characters.");
  for (const step of data.steps) {
    if (step.kind === "split") {
      if (
        step.mode === "delimited" &&
        (![...step.delimiters].length ||
          [...step.delimiters].length > 8 ||
          /["\r\n\0]/.test(step.delimiters))
      )
        throw Error(
          "Choose 1–8 delimiter characters without quotes, line breaks or NUL.",
        );
      if (
        step.mode === "fixed" &&
        (!step.breaks.length ||
          step.breaks.some(
            (breakPosition, index) =>
              index > 0 && breakPosition <= step.breaks[index - 1],
          ))
      )
        throw Error(
          "Fixed-width breaks must be distinct increasing character positions.",
        );
    }
    if (
      step.kind === "convert" &&
      step.to !== "text" &&
      step.decimal === step.group
    )
      throw Error("Decimal and grouping separators must differ.");
    if (
      step.kind === "fill" &&
      new Set(step.groupColumns).size !== step.groupColumns.length
    )
      throw Error("Choose distinct fill group columns.");
    if (
      step.kind === "fill" &&
      step.columns.some((column) => step.groupColumns.includes(column))
    )
      throw Error(
        "Fill targets must differ from the group columns. Use a separate preceding step to prepare group keys.",
      );
  }
  return data;
}
function assertFieldTransformsUnchanged(workbook, updatedWorkbook) {
  if (!workbook.fieldTransforms?.length) return updatedWorkbook;
  const transformsById = new Map(
    workbook.fieldTransforms.map((fieldTransform) => [
      fieldTransform.id,
      fieldTransform,
    ]),
  );
  for (const updatedTransform of updatedWorkbook.fieldTransforms ?? [])
    if (
      transformsById.has(updatedTransform.id) &&
      JSON.stringify(transformsById.get(updatedTransform.id)) !==
        JSON.stringify(updatedTransform)
    )
      throw Error(
        "Saved original field transformation evidence cannot be changed.",
      );
  return {
    ...updatedWorkbook,
    fieldTransforms: [
      ...workbook.fieldTransforms,
      ...(updatedWorkbook.fieldTransforms ?? []).filter(
        (newTransform) => !transformsById.has(newTransform.id),
      ),
    ],
  };
}
const JU = new WeakSet(),
  NC = new WeakSet(),
  YE = new WeakMap(),
  QU = new WeakMap(),
  numberFormatCategories = new Set([
    "general",
    "number",
    "currency",
    "percent",
    "date",
  ]),
  hasOwnKey = (object, key) =>
    Object.prototype.hasOwnProperty.call(object, key),
  requireBoundedText = (value, label, maxLength = 32767) =>
    typeof value == "string" && value.length <= maxLength
      ? value
      : throwWithMessage(
          `${label} must be text of at most ${maxLength} characters.`,
        ),
  P2 = (value, min, max, label) =>
    typeof value == "number" &&
    Number.isInteger(value) &&
    value >= min &&
    value <= max
      ? value
      : throwWithMessage(
          `${label} must be an integer between ${min} and ${max}.`,
        );
function columnIndexToLetters(columnIndex) {
  P2(columnIndex, 0, workbookLimits.maxCols - 1, "Column");
  let letters = "";
  for (
    let oneBasedValue = columnIndex + 1;
    oneBasedValue > 0;
    oneBasedValue = Math.floor((oneBasedValue - 1) / 26)
  )
    letters = String.fromCharCode(65 + ((oneBasedValue - 1) % 26)) + letters;
  return letters;
}
function shiftColumnReference(absoluteMarker, columnLetters, columnDelta) {
  const targetColumnIndex =
    columnLettersToIndex(columnLetters) + (absoluteMarker ? 0 : columnDelta);
  return targetColumnIndex < 0 || targetColumnIndex >= workbookLimits.maxCols
    ? "#REF!"
    : absoluteMarker + columnIndexToLetters(targetColumnIndex);
}
function shiftFormulaReferences(formulaText, rowDelta, columnDelta) {
  if (!Number.isSafeInteger(rowDelta) || !Number.isSafeInteger(columnDelta))
    throw new Error("Formula offsets must be integers.");
  if (!formulaText.startsWith("=")) return formulaText;
  let shiftedFormula = "",
    position = 0;
  for (; position < formulaText.length;) {
    const remainingText = formulaText.slice(position),
      currentChar = formulaText[position],
      sheetPrefixMatch =
        (position === 0 ||
          !unquotedNameCharPattern.test(formulaText[position - 1])) &&
        remainingText.match(sheetReferencePrefixPattern);
    if (sheetPrefixMatch) {
      ((shiftedFormula += sheetPrefixMatch[0]),
        (position += sheetPrefixMatch[0].length));
      continue;
    }
    if (currentChar === '"' || currentChar === "'") {
      const quoteStart = position++;
      for (; position < formulaText.length;)
        if (formulaText[position++] === currentChar)
          if (formulaText[position] === currentChar) position++;
          else break;
      shiftedFormula += formulaText.slice(quoteStart, position);
      continue;
    }
    if (currentChar === "[") {
      const bracketStart = position++;
      let bracketDepth = 1;
      for (; position < formulaText.length && bracketDepth;)
        (formulaText[position] === "[" && bracketDepth++,
          formulaText[position] === "]" && bracketDepth--,
          position++);
      shiftedFormula += formulaText.slice(bracketStart, position);
      continue;
    }
    const isReferenceStart =
        position === 0 ||
        !unquotedNameCharPattern.test(formulaText[position - 1]),
      // A1:B2 is handled as one reference, so that a corner pushed off the
      // grid invalidates the whole range instead of half of it.
      cellRangeMatch = isReferenceStart
        ? /^(\$?)([A-Za-z]{1,3})(\$?)([1-9]\d*):(\$?)([A-Za-z]{1,3})(\$?)([1-9]\d*)/.exec(
            remainingText,
          )
        : null,
      columnRangeMatch = isReferenceStart
        ? /^(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})/.exec(remainingText)
        : null,
      rowRangeMatch = isReferenceStart
        ? /^(\$?)([1-9]\d*):(\$?)([1-9]\d*)/.exec(remainingText)
        : null,
      cellMatch = isReferenceStart
        ? /^(\$?)([A-Za-z]{1,3})(\$?)([1-9]\d*)/.exec(remainingText)
        : null,
      referenceMatch =
        cellRangeMatch ?? columnRangeMatch ?? rowRangeMatch ?? cellMatch;
    if (
      referenceMatch &&
      !unquotedNameCharPattern.test(
        remainingText[referenceMatch[0].length] ?? "",
      ) &&
      !/^\s*[(!]/.test(remainingText.slice(referenceMatch[0].length))
    ) {
      const joinRange = (...parts) =>
        parts.includes("#REF!")
          ? "#REF!"
          : parts.length === 4
            ? `${parts[0]}${parts[1]}:${parts[2]}${parts[3]}`
            : `${parts[0]}:${parts[1]}`;
      if (cellRangeMatch)
        shiftedFormula += joinRange(
          shiftColumnReference(
            cellRangeMatch[1],
            cellRangeMatch[2],
            columnDelta,
          ),
          offsetRowReference(cellRangeMatch[3], cellRangeMatch[4], rowDelta),
          shiftColumnReference(
            cellRangeMatch[5],
            cellRangeMatch[6],
            columnDelta,
          ),
          offsetRowReference(cellRangeMatch[7], cellRangeMatch[8], rowDelta),
        );
      else if (columnRangeMatch)
        shiftedFormula += joinRange(
          shiftColumnReference(
            columnRangeMatch[1],
            columnRangeMatch[2],
            columnDelta,
          ),
          shiftColumnReference(
            columnRangeMatch[3],
            columnRangeMatch[4],
            columnDelta,
          ),
        );
      else if (rowRangeMatch)
        shiftedFormula += joinRange(
          offsetRowReference(rowRangeMatch[1], rowRangeMatch[2], rowDelta),
          offsetRowReference(rowRangeMatch[3], rowRangeMatch[4], rowDelta),
        );
      else if (cellMatch) {
        const shiftedColumn = shiftColumnReference(
            cellMatch[1],
            cellMatch[2],
            columnDelta,
          ),
          shiftedRow = offsetRowReference(cellMatch[3], cellMatch[4], rowDelta);
        shiftedFormula +=
          shiftedColumn === "#REF!" || shiftedRow === "#REF!"
            ? "#REF!"
            : shiftedColumn + shiftedRow;
      }
      position += referenceMatch[0].length;
      continue;
    }
    ((shiftedFormula += currentChar), position++);
  }
  return shiftedFormula;
}

export {
  validateRetainedSourceFile,
  verifyPrintSnapshotGuard,
  AU,
  validateAnalysisReport,
  nonEmptyLabelSchema,
  B2,
  validateDataModelReport,
  requireBoundedText,
  reportOutputFormatVersion,
  sourceImportLimits,
  parseAnalysisRequest,
  requireWhatIfName,
  retainAlternateFormatSource,
  validatePrintPageBreaks,
  headerFooterTokenPattern,
  textJoinSpecSchema,
  daxFunctionNames,
  HO,
  alternateFormatExtensions,
  readRetainedXlsxSource,
  transformCellValueSchema,
  JU,
  validateDelimitedSourceFootprint,
  resolvePageDimensions,
  validateFieldTransformRecipe,
  walkDaxSyntaxTree,
  assertPlainObjectKeys,
  DaxError,
  whatIfLimits,
  NC,
  P2,
  columnarModelFormatVersion,
  expandWhatIfScenarios,
  QU,
  referenceImpactKinds,
  createBoundedTextSchema,
  parseDaxExpression,
  validateModelQuery,
  xlsxInspectionLimits,
  textCleanupFunctionNames,
  validateSourceImports,
  cellAddressKey,
  reportOutputLimits,
  hashRecordExcludingHash,
  validateDataModel,
  requireRetainedXlsxSource,
  WU,
  decodeAutomationArtifact,
  defaultPrintHeaderFooter,
  exportHeaderFooterTemplate,
  requireAlternateFormatOriginal,
  YE,
  validateTextCleanupSteps,
  validateWhatIfFields,
  validateAutomationArtifacts,
  textInferenceLimits,
  sourceRowValueSchema,
  encodeBytesToBase64,
  assertFieldTransformsUnchanged,
  isBoundedPlainText,
  validateWhatIfValues,
  throwTextCleanupEvidenceError,
  analysisLimits,
  analysisFormatVersion,
  jm,
  validateExtractedSourceTables,
  requireWhatIfList,
  assertTextCleanupEvidenceUnchanged,
  automationLimits,
  throwDaxError,
  normalizeAlternateFormatSource,
  expandHeaderFooterTemplate,
  daxExpressionLimits,
  dataModelLimits,
  columnIndexToLetters,
  hasOwnKey,
  printRenderLimits,
  shiftFormulaReferences,
  numberFormatCategories,
  reportHashSchema,
  resolveExtractedCellValue,
  buildRetainedSourceRecord,
};

/* Also used by the workers. */
export {
  IU,
  Poe,
  analysisCellValueSchema,
  analysisOutcomeSchema,
  analysisReportRecordSchema,
  analysisSpecSchema,
  boundedNumberSchema,
  cellReferenceSchema,
  cellTextValueSchema,
  columnIndexSchema,
  dataModelDefinitionSchema,
  dataModelReportSchema,
  distinctColumnPositions,
  distinctExampleRowsSchema,
  documentSourceFormats,
  extractedDocumentSchema,
  extractedTableSchema,
  fieldTransformRecipeSchema,
  fieldTransformStepSchema,
  finiteNumberSchema,
  modelColumnSchema,
  modelFilterSchema,
  modelGroupFieldSchema,
  modelMeasureSchema,
  modelQuerySchema,
  modelTableSourceSchema,
  nativeCellValueSchema,
  outputWorksheetNameSchema,
  relationshipOptionsSchema,
  retainedSourceFileSchema,
  reviewCorrectionSchema,
  shiftColumnReference,
  solverAnalysisSpecSchema,
  sourceImportRecordSchema,
  sourceRegionSchema,
  statisticsAnalysisSpecSchema,
  tableIdentifierSchema,
  textCaseModeSchema,
  textExtractStepSchema,
};

/* Also used by the workers. */
export {
  MU,
  assertUniqueCaseInsensitive,
  automationArtifactKinds,
  isSafeFileName,
  requireAutomationEvidenceCount,
  requireAutomationFileName,
  requireAutomationSourceHash,
  requireAutomationText,
  requireAutomationTimestamp,
  retainOriginalXlsx,
  validateAlternateSource,
  validateAutomationApprovedRange,
  validateRetainedXlsxSource,
};
