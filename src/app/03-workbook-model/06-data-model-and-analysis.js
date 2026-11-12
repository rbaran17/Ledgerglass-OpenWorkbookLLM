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
      (e) => !/[\\/?*\[\]:\u0000-\u001f]/.test(e),
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
  const r = analysisSpecSchema.safeParse(request);
  if (!r.success)
    throw Error(
      `Invalid analysis request: ${r.error.issues[0]?.path.join(".")}: ${r.error.issues[0]?.message}`,
    );
  const data = r.data;
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
function validateAnalysisReport(e, workbook) {
  const t = e;
  let n = 0;
  for (const sheet of workbook.sheets) {
    const l = t.sheets.find((m) => m.id === sheet.id)?.analysisReport;
    if (l === undefined) continue;
    if (
      ((n += JSON.stringify(l).length), n > analysisLimits.metadataCharacters)
    )
      throw Error("Analysis metadata exceeds 2,000,000 characters.");
    const c = analysisReportRecordSchema.safeParse(l);
    if (!c.success)
      throw Error(`Invalid analysis report: ${c.error.issues[0]?.message}`);
    const data = c.data;
    if (
      ((data.request = parseAnalysisRequest(data.request)),
      data.definitionHash !== fingerprintCanonicalValue(data.request) ||
        !data.result.success ||
        !["complete", "optimal"].includes(data.result.status) ||
        data.result.rows.reduce((m, v) => m + v.length, 0) >
          analysisLimits.outputCells)
    )
      throw Error("Analysis report metadata is inconsistent.");
    const f = /^([A-Z]{1,3})([1-9]\d{0,6}):([A-Z]{1,3})([1-9]\d{0,6})$/.exec(
        data.outputRange,
      ),
      p = (m) => [...m].reduce((v, g) => v * 26 + g.charCodeAt(0) - 64, 0);
    if (
      !f ||
      p(f[3]) > 16384 ||
      Number(f[4]) > 1048576 ||
      p(f[3]) < p(f[1]) ||
      Number(f[4]) < Number(f[2]) ||
      (p(f[3]) - p(f[1]) + 1) * (Number(f[4]) - Number(f[2]) + 1) >
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
  createBoundedTextSchema = (e) => string().max(e),
  nonEmptyLabelSchema = string().min(1).max(200),
  reportHashSchema = string().regex(/^[a-f0-9]{64}$/),
  AU = string().datetime();
const hashRecordExcludingHash = (e) => {
  const { hash, ...t } = e;
  return fingerprintCanonicalValue(t);
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
  constructor(r, t) {
    (super(t), (this.code = r), (this.name = "DaxError"));
  }
}
const throwDaxError = (e, r) => {
  throw new DaxError(e, `DAX: ${r}`);
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
  const r = expression.trim().replace(/^=/, ""),
    t = [];
  let position = 0;
  const o = (m, v, g) => {
    const b = position++;
    let y = "",
      w = false;
    for (; position < r.length;) {
      const E = r[position++];
      if (E === v)
        if (r[position] === v) ((y += v), position++);
        else {
          w = true;
          break;
        }
      else y += E;
    }
    (w || throwDaxError("DAX_SYNTAX", `Unclosed ${m} at position ${b + 1}.`),
      t.push({ kind: g, value: y, position: b }));
  };
  for (; position < r.length;) {
    if (/\s/.test(r[position])) {
      position++;
      continue;
    }
    if (r[position] === '"') {
      o('"', '"', "string");
      continue;
    }
    if (r[position] === "'") {
      o("'", "'", "quoted");
      continue;
    }
    if (r[position] === "[") {
      o("[", "]", "bracket");
      continue;
    }
    const m = r.slice(position),
      v = /^(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?/.exec(m),
      g = /^[A-Za-z_][A-Za-z0-9_]*/.exec(m),
      b = /^(?:==|<>|<=|>=|&&|\|\||[+\-*/^&=<>(),])/.exec(m),
      y = v ?? g ?? b;
    (y ||
      throwDaxError(
        "DAX_UNSUPPORTED",
        `Unsupported syntax at position ${position + 1}. Use the documented DAX subset; worksheet and query syntax are not evaluated.`,
      ),
      t.push({
        kind: v ? "number" : g ? "identifier" : "op",
        value: y[0],
        position: position,
      }),
      (position += y[0].length),
      t.length > daxExpressionLimits.syntaxNodes * 3 &&
        throwDaxError("DAX_LIMIT", "Expression token budget exceeded."));
  }
  t.push({ kind: "end", value: "", position: position });
  let a = 0,
    s = 0;
  const d = () => t[a],
    u = () => t[a++],
    l = (m) => {
      u().value !== m && throwDaxError("DAX_SYNTAX", `Expected ${m}.`);
    },
    c = (m) => (
      ++s > daxExpressionLimits.syntaxNodes &&
        throwDaxError("DAX_LIMIT", "Expression exceeds 512 syntax nodes."),
      m
    ),
    h = {
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
    f = (m = 0, v = 0) => {
      v > daxExpressionLimits.depth &&
        throwDaxError("DAX_LIMIT", "Expression nesting exceeds 48 levels.");
      const g = u();
      let left;
      if (g.kind === "number") {
        const y = Number(g.value);
        (Number.isFinite(y) ||
          throwDaxError(
            "DAX_NUMBER",
            "Numeric literal exceeds finite precision.",
          ),
          (left = c({ kind: "literal", value: y })));
      } else if (g.kind === "string")
        left = c({ kind: "literal", value: g.value });
      else if (g.value === "(") ((left = f(0, v + 1)), l(")"));
      else if (
        g.value === "+" ||
        g.value === "-" ||
        (g.kind === "identifier" && g.value.toUpperCase() === "NOT")
      )
        left = c({
          kind: "unary",
          op: g.value.toUpperCase(),
          value: f(g.value.toUpperCase() === "NOT" ? 3 : 8, v + 1),
        });
      else if (g.kind === "bracket")
        left = c({ kind: "measure", name: g.value });
      else if (g.kind === "identifier" || g.kind === "quoted")
        if (d().value === "(" && g.kind === "identifier") {
          const y = g.value.toUpperCase();
          (Object.prototype.hasOwnProperty.call(daxFunctionNames, y) ||
            throwDaxError(
              "DAX_UNSUPPORTED",
              `Unsupported function ${g.value}. No worksheet fallback is used.`,
            ),
            u());
          const args = [];
          if (d().value !== ")")
            do {
              if ((args.push(f(0, v + 1)), d().value !== ",")) break;
              u();
            } while (true);
          l(")");
          const E = daxFunctionNames[y];
          ((args.length < E[0] || args.length > E[1]) &&
            throwDaxError(
              "DAX_ARGUMENT",
              `${y} expects ${E[0] === E[1] ? E[0] : `${E[0]}–${E[1]}`} arguments in this subset.`,
            ),
            (left = c({ kind: "call", name: y, args: args })));
        } else
          d().kind === "bracket"
            ? (left = c({ kind: "field", table: g.value, column: u().value }))
            : g.kind === "identifier" &&
                ["TRUE", "FALSE"].includes(g.value.toUpperCase())
              ? (left = c({
                  kind: "literal",
                  value: g.value.toUpperCase() === "TRUE",
                }))
              : (left = c({ kind: "table", name: g.value }));
      else
        return throwDaxError(
          "DAX_SYNTAX",
          `Expected an expression at position ${g.position + 1}.`,
        );
      for (
        ;
        d().kind === "op" && h[d().value] !== undefined && h[d().value] >= m;
      ) {
        const value = u().value,
          w = h[value];
        left = c({
          kind: "binary",
          op: value,
          left: left,
          right: f(w + 1, v + 1),
        });
      }
      return left;
    },
    p = f();
  return (
    d().kind !== "end" &&
      throwDaxError(
        "DAX_UNSUPPORTED",
        `Unexpected ${d().value}. VAR, RETURN, EVALUATE, worksheet references and unsupported syntax are not evaluated.`,
      ),
    p
  );
}
function walkDaxSyntaxTree(e, r) {
  (r(e),
    e.kind === "call"
      ? e.args.forEach((arg) => walkDaxSyntaxTree(arg, r))
      : e.kind === "binary"
        ? (walkDaxSyntaxTree(e.left, r), walkDaxSyntaxTree(e.right, r))
        : e.kind === "unary" && walkDaxSyntaxTree(e.value, r));
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
    .refine((e) => !/[\u0000-\u001f\u007f]/.test(e)),
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
  assertUniqueCaseInsensitive = (e, r) => {
    if (new Set(e.map((t) => t.toLowerCase())).size !== e.length)
      throw Error(`${r} must be unique, including case variants.`);
  };
function validateDataModel(e) {
  const r = dataModelDefinitionSchema.safeParse(e);
  if (!r.success)
    throw Error(
      `Invalid data model: ${r.error.issues[0]?.path.join(".")}: ${r.error.issues[0]?.message}`,
    );
  const data = r.data;
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
      (s, d) =>
        s +
        d.columns.filter((column) => column.expression !== undefined).length,
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
  const n = (s) => {
    const table = data.tables.find((table) => table.id === s.tableId),
      u = table?.columns.find((column) => column.name === s.column);
    if (!u) throw Error(`Unknown model field ${s.tableId}.${s.column}.`);
    return u;
  };
  for (const relationship of data.relationships) {
    if (relationship.one.tableId === relationship.many.tableId)
      throw Error(
        "Self relationships are unsupported; use distinct role-playing tables.",
      );
    if (n(relationship.one).type !== n(relationship.many).type)
      throw Error(
        `Relationship ${relationship.id} requires matching key types.`,
      );
    if (
      n(relationship.one).expression !== undefined ||
      n(relationship.many).expression !== undefined
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
        const d = n({ tableId: measure.tableId, column: measure.column ?? "" });
        if (
          ["sum", "average", "min", "max"].includes(measure.operation) &&
          d.type !== "number"
        )
          throw Error(`Measure ${measure.name} requires a number column.`);
      }
    }
  }
  return data;
}
function validateModelQuery(query, definition) {
  const t = modelQuerySchema.safeParse(query);
  if (!t.success)
    throw Error(`Invalid model query: ${t.error.issues[0]?.message}`);
  const data = t.data;
  (assertUniqueCaseInsensitive(data.measureIds, "Selected measures"),
    assertUniqueCaseInsensitive(
      data.groupBy.map((s) => JSON.stringify(s)),
      "Group fields",
    ));
  for (const s of [...data.groupBy, ...data.filters])
    if (
      !definition.tables
        .find((table) => table.id === s.tableId)
        ?.columns.some((column) => column.name === s.column)
    )
      throw Error(`Unknown query field ${s.tableId}.${s.column}.`);
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
function validateDataModelReport(e, workbook) {
  const t = e;
  let n = 0;
  if (t.dataModel !== undefined) {
    if (
      ((n += JSON.stringify(t.dataModel).length),
      n > dataModelLimits.metadataCharacters)
    )
      throw Error("Data model metadata exceeds 2,000,000 characters.");
    workbook.dataModel = validateDataModel(t.dataModel);
  }
  for (const sheet of workbook.sheets) {
    const d = t.sheets.find((c) => c.id === sheet.id)?.dataModelReport;
    if (d === undefined) continue;
    if (
      ((n += JSON.stringify(d).length), n > dataModelLimits.metadataCharacters)
    )
      throw Error("Data model metadata exceeds 2,000,000 characters.");
    const u = dataModelReportSchema.safeParse(d);
    if (!u.success)
      throw Error(`Invalid model report: ${u.error.issues[0]?.message}`);
    const data = u.data;
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
      (e) => e.x + e.width <= 1.000001 && e.y + e.height <= 1.000001,
      "Source region must fit its page.",
    );
function Poe(e, r) {
  if (typeof e != "string") return false;
  try {
    const t = JSON.parse(e);
    return r === "array"
      ? Array.isArray(t)
      : !!t && typeof t == "object" && !Array.isArray(t);
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
      (e) =>
        e.type === "array" || e.type === "object"
          ? Poe(e.value, e.type)
          : e.type === "null"
            ? e.value === null
            : typeof e.value === e.type,
      "Native value and type disagree.",
    )
    .refine(
      (e) => !e.missing || (e.value === null && e.type === "null"),
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
    .superRefine((e, r) => {
      (new Set(e.columns.map((column) => column.path)).size !==
        e.columns.length ||
        e.rows.some((row) => row.length !== e.columns.length)) &&
        r.addIssue({
          code: "custom",
          message: "Table columns and rows must align with unique paths.",
        });
    }),
  retainedSourceFileSchema = object({
    fileName: createBoundedTextSchema(255)
      .min(1)
      .refine((e) => !/[\\/\x00-\x1f\x7f\u202a-\u202e\u2066-\u2069]/.test(e)),
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
      (e) => e.decision !== "correct" || e.value !== undefined,
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
function validateExtractedSourceTables(e) {
  const r = extractedDocumentSchema.parse(e);
  let t = 0;
  if (new Set(r.tables.map((table) => table.id)).size !== r.tables.length)
    throw Error("Source table IDs must be unique.");
  for (const table of r.tables)
    for (const row of table.rows)
      for (const a of row) {
        if (++t > sourceImportLimits.cells)
          throw Error("Source cell budget exceeded.");
        if (
          r.method === "document-adapter" &&
          (!a.region ||
            !r.pageCount ||
            a.region.page > r.pageCount ||
            (["png", "jpeg"].includes(r.format) && r.pageCount !== 1))
        )
          throw Error(
            "Document cells require a source page and bounding region.",
          );
      }
  if (
    (r.format === "json" && r.method !== "native-json") ||
    (r.format === "xml" && r.method !== "native-xml") ||
    (!["json", "xml"].includes(r.format) && r.method !== "document-adapter")
  )
    throw Error("Source format and extraction method disagree.");
  return r;
}
function validateRetainedSourceFile(original) {
  const r = retainedSourceFileSchema.parse(original),
    t = r.fileName.split(".").at(-1)?.toLowerCase();
  if ((t === "jpg" ? "jpeg" : t === "htm" ? "html" : t) !== r.format)
    throw Error("Source extension and format disagree.");
  if (
    r.base64.length !== Math.ceil(r.byteLength / 3) * 4 ||
    !/^[A-Za-z0-9+/]*={0,2}$/.test(r.base64)
  )
    throw Error("Invalid retained source encoding.");
  const n = atob(r.base64);
  if (btoa(n) !== r.base64 || n.length !== r.byteLength)
    throw Error("Retained source byte length changed.");
  const o = Uint8Array.from(n, (s) => s.charCodeAt(0));
  if (sha256HexDigest(o) !== r.sha256)
    throw Error("Retained source checksum mismatch.");
  return o;
}
function buildRetainedSourceRecord(e, fileName, format) {
  let n = "";
  for (let a = 0; a < e.length; a += 8192)
    n += String.fromCharCode(...e.subarray(a, a + 8192));
  const original = {
    fileName: fileName,
    format: format,
    byteLength: e.length,
    sha256: sha256HexDigest(e),
    base64: btoa(n),
  };
  return (validateRetainedSourceFile(original), original);
}
function resolveExtractedCellValue(e, r) {
  if (e.confidence < sourceImportLimits.confidence && !r)
    throw Error("Review every low-confidence cell before importing.");
  if (!r) return e.value;
  const t = reviewCorrectionSchema.parse(r);
  return t.decision === "exclude"
    ? null
    : t.decision === "correct"
      ? t.value
      : e.value;
}
const uk = new WeakSet();
function MU(e) {
  Object.freeze(e);
  for (const r of Object.values(e)) r && typeof r == "object" && MU(r);
}
function validateSourceImports(workbook, r) {
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
  let t = 0;
  const n = new Set();
  if (
    ((r.sourceImports = workbook.sourceImports.map((sourceImport) => {
      if (
        sourceImport &&
        typeof sourceImport == "object" &&
        uk.has(sourceImport)
      ) {
        const d = sourceImport;
        if (((t += d.original.byteLength), n.has(d.id)))
          throw Error("Duplicate source import ID.");
        return (n.add(d.id), d);
      }
      const a = sourceImportRecordSchema.parse(sourceImport);
      if (
        (validateRetainedSourceFile(a.original),
        validateExtractedSourceTables({
          schemaVersion: 1,
          fileName: a.original.fileName,
          format: a.original.format,
          byteLength: a.original.byteLength,
          sha256: a.original.sha256,
          tables: [a.table],
          method: a.method,
          extractor: a.extractor,
          ...(a.pageCount ? { pageCount: a.pageCount } : {}),
          warnings: [],
        }),
        a.table.ambiguous && !a.ambiguityReviewed)
      )
        throw Error("Ambiguous source tables require review.");
      const s = new Set();
      if (
        (a.table.rows.forEach((row, u) =>
          row.forEach((l, c) => {
            const h = `${u}:${c}`;
            (s.add(h), resolveExtractedCellValue(l, a.resolutions[h]));
          }),
        ),
        Object.keys(a.resolutions).some((resolution) => !s.has(resolution)))
      )
        throw Error("Unknown source review cell.");
      if (n.has(a.id)) throw Error("Duplicate source import ID.");
      return (n.add(a.id), (t += a.original.byteLength), MU(a), uk.add(a), a);
    })),
    t > sourceImportLimits.retainedBytes)
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
  let r;
  try {
    r = atob(xlsxSource.base64);
  } catch {
    throw Error("Invalid retained XLSX source encoding.");
  }
  if (btoa(r) !== xlsxSource.base64 || r.length !== xlsxSource.byteLength)
    throw Error("Retained XLSX source byte length or encoding changed.");
  const t = Uint8Array.from(r, (n) => n.charCodeAt(0));
  if (sha256HexDigest(t) !== xlsxSource.sha256)
    throw Error("Retained XLSX source checksum does not match its bytes.");
  return t;
}
function retainOriginalXlsx(e, fileName) {
  if (e.byteLength > xlsxInspectionLimits.maxBytes)
    throw Error("Original XLSX exceeds the 10 MiB retention limit.");
  let t = "";
  for (let o = 0; o < e.length; o += 8192)
    t += String.fromCharCode(...e.subarray(o, o + 8192));
  const xlsxSource = {
    schemaVersion: 1,
    fileName: fileName,
    byteLength: e.length,
    sha256: sha256HexDigest(e),
    base64: btoa(t),
  };
  return (
    validateRetainedXlsxSource(xlsxSource),
    Object.freeze(xlsxSource),
    q5.add(xlsxSource),
    xlsxSource
  );
}
function readRetainedXlsxSource(workbook, r) {
  if (workbook.xlsxSource === undefined) return;
  const xlsxSource = workbook.xlsxSource;
  if (xlsxSource && typeof xlsxSource == "object" && q5.has(xlsxSource)) {
    r.xlsxSource = xlsxSource;
    return;
  }
  (validateRetainedXlsxSource(xlsxSource),
    (r.xlsxSource = Object.freeze({
      schemaVersion: 1,
      fileName: xlsxSource.fileName,
      byteLength: xlsxSource.byteLength,
      sha256: xlsxSource.sha256,
      base64: xlsxSource.base64,
    })),
    q5.add(r.xlsxSource));
}
function requireRetainedXlsxSource(e) {
  if (!e.xlsxSource)
    throw Error(
      "This workbook has no retained XLSX source. Open an XLSX file to retain its original bytes.",
    );
  return {
    bytes: validateRetainedXlsxSource(e.xlsxSource),
    fileName: e.xlsxSource.fileName,
    sha256: e.xlsxSource.sha256,
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
  isSafeFileName = (e) =>
    typeof e == "string" &&
    e.length > 0 &&
    e.length <= 255 &&
    !/[\\/\x00-\x1f\x7f\u202a-\u202e\u2066-\u2069]/.test(e),
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
function retainAlternateFormatSource(e, fileName, format, n) {
  const alternateSource = {
    ...retainOriginalXlsx(e, "retained.xlsx"),
    fileName: fileName,
    format: format,
    conversion: { ...n },
  };
  return (
    validateAlternateSource(alternateSource),
    Object.freeze(alternateSource.conversion),
    Object.freeze(alternateSource),
    jU.add(alternateSource),
    alternateSource
  );
}
function normalizeAlternateFormatSource(workbook, r) {
  if (workbook.alternateSource === undefined) return;
  const alternateSource = workbook.alternateSource;
  if (jU.has(alternateSource)) {
    r.alternateSource = alternateSource;
    return;
  }
  const n = validateAlternateSource(alternateSource);
  r.alternateSource = retainAlternateFormatSource(
    n,
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
function encodeBytesToBase64(e) {
  let r = "";
  for (let t = 0; t < e.length; t += 8192)
    r += String.fromCharCode(...e.subarray(t, t + 8192));
  return btoa(r);
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
  const r = atob(base64);
  if (btoa(r) !== base64)
    throw Error("Automation artifact is not canonical base64.");
  return Uint8Array.from(r, (t) => t.charCodeAt(0));
}
const requireAutomationText = (e, r, t) => {
    if (typeof e != "string" || !e.trim() || e.length > r)
      throw Error(`Invalid automation ${t}.`);
    return e;
  },
  requireAutomationSourceHash = (e) => {
    if (typeof e != "string" || !/^[a-f0-9]{64}$/.test(e))
      throw Error("Invalid automation source hash.");
    return e;
  },
  requireAutomationEvidenceCount = (e, r, t) => {
    if (!Number.isSafeInteger(e) || Number(e) < r || Number(e) > t)
      throw Error("Invalid automation evidence count.");
    return Number(e);
  },
  requireAutomationTimestamp = (e) => {
    const r = requireAutomationText(e, 40, "timestamp");
    if (!/^\d{4}-\d\d-\d\dT/.test(r) || !Number.isFinite(Date.parse(r)))
      throw Error("Invalid automation timestamp.");
    return r;
  },
  validateAutomationApprovedRange = (e) => {
    const r = requireAutomationText(e, 40, "approved range"),
      t = /^([A-Z]{1,3})([1-9]\d*):([A-Z]{1,3})([1-9]\d*)$/.exec(r);
    if (!t) throw Error("Invalid automation approved range.");
    const n = (u) => [...u].reduce((l, c) => l * 26 + c.charCodeAt(0) - 64, 0),
      [o, a, s, d] = [n(t[1]), Number(t[2]), n(t[3]), Number(t[4])];
    if (
      o > s ||
      a > d ||
      s > 16384 ||
      d > 1048576 ||
      (s - o + 1) * (d - a + 1) > automationLimits.maxCells
    )
      throw Error("Invalid automation approved range.");
    return r;
  };
function requireAutomationFileName(e) {
  const r = requireAutomationText(e, 255, "file name");
  if (/[\\/\x00-\x1f\x7f]/.test(r) || r === "." || r === "..")
    throw Error("Invalid automation file name.");
  return r;
}
function validateAutomationArtifacts(workbook, r) {
  if (workbook.automationArtifacts !== undefined) {
    if (
      !Array.isArray(workbook.automationArtifacts) ||
      workbook.automationArtifacts.length > automationLimits.maxArtifacts
    )
      throw Error("Keep at most 20 automation artifacts.");
    let t = 0;
    const n = new Set();
    r.automationArtifacts = workbook.automationArtifacts.map(
      (automationArtifact) => {
        if (
          !automationArtifact ||
          typeof automationArtifact != "object" ||
          Array.isArray(automationArtifact)
        )
          throw Error("Invalid automation artifact.");
        const a = automationArtifact,
          s = requireAutomationText(a.id, 100, "artifact ID");
        if (n.has(s)) throw Error("Duplicate automation artifact ID.");
        n.add(s);
        const d = decodeAutomationArtifact(a.base64);
        if (
          ((t += d.length),
          !d.length ||
            d.length !== a.byteLength ||
            d.length > automationLimits.maxArtifactBytes ||
            t > automationLimits.maxStoredBytes)
        )
          throw Error(
            "Automation artifacts exceed their 2 MB per-file or 10 MB retained-byte budget.",
          );
        if (sha256HexDigest(d) !== requireAutomationSourceHash(a.sha256))
          throw Error(
            "Automation artifact checksum does not match its preserved bytes.",
          );
        if (!automationArtifactKinds.includes(a.kind))
          throw Error("Unknown automation artifact kind.");
        return {
          id: s,
          fileName: requireAutomationFileName(a.fileName),
          kind: a.kind,
          sha256: a.sha256,
          byteLength: d.length,
          base64: a.base64,
          retainedAt: requireAutomationTimestamp(a.retainedAt),
          retainedBy: requireAutomationText(a.retainedBy, 100, "retainer"),
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
    const t = new Set();
    r.automationRuns = workbook.automationRuns.map((automationRun) => {
      if (
        !automationRun ||
        typeof automationRun != "object" ||
        Array.isArray(automationRun)
      )
        throw Error("Invalid automation run evidence.");
      const o = automationRun,
        a = requireAutomationText(o.id, 100, "run ID");
      if (t.has(a)) throw Error("Duplicate automation run ID.");
      if (
        (t.add(a),
        !["vba", "office-script"].includes(o.language) ||
          o.adapter !== "literal-ranges-v1")
      )
        throw Error("Unknown automation migration adapter.");
      if (
        !Array.isArray(o.sourceLines) ||
        o.sourceLines.length > automationLimits.maxCommands
      )
        throw Error("Invalid automation source lines.");
      const beforeRevision = requireAutomationEvidenceCount(
          o.beforeRevision,
          0,
          Number.MAX_SAFE_INTEGER - 2,
        ),
        afterRevision = requireAutomationEvidenceCount(
          o.afterRevision,
          1,
          Number.MAX_SAFE_INTEGER - 1,
        );
      if (afterRevision !== beforeRevision + 1)
        throw Error("Automation run revisions are inconsistent.");
      if (o.sourceLines.length !== o.commands)
        throw Error(
          "Automation source-line evidence must match the command count.",
        );
      return {
        id: a,
        eventId: requireAutomationText(o.eventId, 100, "event ID"),
        sourceName: requireAutomationFileName(o.sourceName),
        sourceSha256: requireAutomationSourceHash(o.sourceSha256),
        codeSha256: requireAutomationSourceHash(o.codeSha256),
        language: o.language,
        adapter: o.adapter,
        actor: requireAutomationText(o.actor, 100, "actor"),
        timestamp: requireAutomationTimestamp(o.timestamp),
        beforeRevision: beforeRevision,
        afterRevision: afterRevision,
        sheetId: requireAutomationText(o.sheetId, 100, "sheet ID"),
        range: validateAutomationApprovedRange(o.range),
        commands: requireAutomationEvidenceCount(
          o.commands,
          1,
          automationLimits.maxCommands,
        ),
        changedCells: requireAutomationEvidenceCount(
          o.changedCells,
          0,
          automationLimits.maxCells,
        ),
        sourceLines: o.sourceLines.map((sourceLine) =>
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
function validateDelimitedSourceFootprint(delimitedSource, r) {
  if (delimitedSource === undefined) return;
  if (
    !delimitedSource ||
    typeof delimitedSource != "object" ||
    Array.isArray(delimitedSource)
  )
    throw Error("Invalid delimited source footprint.");
  const t = delimitedSource;
  if (
    !Number.isInteger(t.records) ||
    Number(t.records) < 1 ||
    Number(t.records) > workbookLimits.maxRows ||
    !Number.isInteger(t.columns) ||
    Number(t.columns) < 1 ||
    Number(t.columns) > workbookLimits.maxCols ||
    typeof t.hasHeader != "boolean"
  )
    throw Error(
      "Delimited source footprint exceeds worksheet bounds or has invalid header metadata.",
    );
  r.delimitedSource = {
    records: Number(t.records),
    columns: Number(t.columns),
    hasHeader: t.hasHeader,
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
  HO = (e, r) => Object.keys(e).every((t) => r.includes(t)),
  isBoundedPlainText = (e, r) =>
    typeof e == "string" &&
    !!e.trim() &&
    e.length <= r &&
    !/[\u0000-\u001f]/.test(e);
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
  cellAddressKey = (e) => JSON.stringify([e.sheetId, e.address]);
function validateWhatIfFields(e, r) {
  if (
    !e ||
    typeof e != "object" ||
    Array.isArray(e) ||
    ![Object.prototype, null].includes(Object.getPrototypeOf(e)) ||
    Object.keys(e).some((t) => !r.includes(t))
  )
    throw Error("Invalid what-if analysis fields.");
  return e;
}
function requireWhatIfName(e, r = 100) {
  if (
    typeof e != "string" ||
    !e.trim() ||
    e.length > r ||
    /[\u0000-\u001f]/.test(e)
  )
    throw Error(
      `What-if names must be nonempty text of at most ${r} characters.`,
    );
  return e.trim();
}
function requireWhatIfList(e, r, t = 1) {
  if (!Array.isArray(e) || e.length < t || e.length > r)
    throw Error(`What-if lists require ${t}–${r} entries.`);
  return Array.from(e);
}
function validateWhatIfValues(e, r) {
  return requireWhatIfList(e, r).map((t) => {
    if (!isFiniteNumber(t))
      throw Error("Every what-if input value must be a finite number.");
    return t;
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
function resolvePrintIdentityToken(e, workbook) {
  return e === "Workbook"
    ? (workbook?.name ?? "[Workbook not available]")
    : e === "Revision"
      ? workbook
        ? String(workbook.revision)
        : "[Revision not available]"
      : e === "Engagement"
        ? workbook?.metadata?.engagement || "[Engagement not set]"
        : workbook?.metadata?.period || "[Period not set]";
}
function expandHeaderFooterTemplate(e, workbook, sheet, n, total) {
  return e.replace(headerFooterTokenPattern, (a, s) =>
    s
      ? resolvePrintIdentityToken(s, workbook)
      : a === "&&"
        ? "&"
        : a === "&A"
          ? sheet.name
          : String(a === "&P" ? n : total),
  );
}
function exportHeaderFooterTemplate(e, workbookInput) {
  const t = e.replace(headerFooterTokenPattern, (n, o) => {
    if (!o) return n;
    if (!workbookInput)
      throw Error(
        "Exporting print identity fields requires the current workbook snapshot.",
      );
    return resolvePrintIdentityToken(o, workbookInput).replace(/&/g, "&&");
  });
  if (t.length > 255)
    throw Error(
      "Resolved print identifiers exceed Excel's 255-character header/footer limit. Shorten the header, footer or workbook properties before exporting.",
    );
  return t;
}
function validatePrintPageBreaks(e, r) {
  if (e !== undefined) {
    if (
      !Array.isArray(e) ||
      e.length > printRenderLimits.breaksPerAxis ||
      e.some(
        (t, n) =>
          !Number.isSafeInteger(t) ||
          t < 1 ||
          t >= r ||
          (n > 0 && t <= e[n - 1]),
      )
    )
      throw Error(
        "Manual page breaks require at most 100 distinct ascending boundaries within each worksheet axis (before row/column 2 or later).",
      );
    return e.length ? [...e] : undefined;
  }
}

function verifyPrintSnapshotGuard(workbook, r) {
  if (
    r.workbookId !== workbook.id ||
    r.revision !== workbook.revision ||
    r.snapshotGuard !== JSON.stringify(workbook)
  )
    throw Error(
      "The workbook changed after page setup opened. Close and reopen page setup.",
    );
}
function resolvePageDimensions(e) {
  const r =
    e.paper === "a4"
      ? [793.7007874015749, 1122.5196850393702]
      : e.paper === "legal"
        ? [816, 1344]
        : [816, 1056];
  return e.orientation === "landscape"
    ? { width: r[1], height: r[0] }
    : { width: r[0], height: r[1] };
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
function assertPlainObjectKeys(e, r) {
  (!isPlainObjectValue(e) || Object.keys(e).some((t) => !r.includes(t))) &&
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
function jm(e, r) {
  return typeof e == "string" && e.length > 0 && e.length <= r;
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
function assertTextCleanupEvidenceUnchanged(workbook, r) {
  if (!workbook.textCleanups?.length) return r;
  const t = new Map(
    workbook.textCleanups.map((textCleanup) => [textCleanup.id, textCleanup]),
  );
  for (const o of r.textCleanups ?? []) {
    const a = t.get(o.id);
    if (a && JSON.stringify(a) !== JSON.stringify(o))
      throw Error("Saved original text cleanup evidence cannot be changed.");
  }
  return {
    ...r,
    textCleanups: [
      ...workbook.textCleanups,
      ...(r.textCleanups ?? []).filter((o) => !t.has(o.id)),
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
      (e) => new Set(e).size === e.length,
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
      (e) => new Set(e.map((r) => r.row)).size === e.length,
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
  const r = fieldTransformRecipeSchema.safeParse(request);
  if (!r.success)
    throw Error(
      `Invalid field transformation: ${r.error.issues[0]?.message ?? "review the recipe"}`,
    );
  const data = r.data;
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
          step.breaks.some((a, s) => s > 0 && a <= step.breaks[s - 1]))
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
function assertFieldTransformsUnchanged(workbook, r) {
  if (!workbook.fieldTransforms?.length) return r;
  const t = new Map(
    workbook.fieldTransforms.map((fieldTransform) => [
      fieldTransform.id,
      fieldTransform,
    ]),
  );
  for (const o of r.fieldTransforms ?? [])
    if (t.has(o.id) && JSON.stringify(t.get(o.id)) !== JSON.stringify(o))
      throw Error(
        "Saved original field transformation evidence cannot be changed.",
      );
  return {
    ...r,
    fieldTransforms: [
      ...workbook.fieldTransforms,
      ...(r.fieldTransforms ?? []).filter((o) => !t.has(o.id)),
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
  hasOwnKey = (e, r) => Object.prototype.hasOwnProperty.call(e, r),
  requireBoundedText = (e, r, t = 32767) =>
    typeof e == "string" && e.length <= t
      ? e
      : throwWithMessage(`${r} must be text of at most ${t} characters.`),
  P2 = (e, r, t, n) =>
    typeof e == "number" && Number.isInteger(e) && e >= r && e <= t
      ? e
      : throwWithMessage(`${n} must be an integer between ${r} and ${t}.`);
function columnIndexToLetters(e) {
  P2(e, 0, workbookLimits.maxCols - 1, "Column");
  let r = "";
  for (let t = e + 1; t > 0; t = Math.floor((t - 1) / 26))
    r = String.fromCharCode(65 + ((t - 1) % 26)) + r;
  return r;
}
function shiftColumnReference(e, r, t) {
  const n = columnLettersToIndex(r) + (e ? 0 : t);
  return n < 0 || n >= workbookLimits.maxCols
    ? "#REF!"
    : e + columnIndexToLetters(n);
}
function shiftFormulaReferences(e, r, t) {
  if (!Number.isSafeInteger(r) || !Number.isSafeInteger(t))
    throw new Error("Formula offsets must be integers.");
  if (!e.startsWith("=")) return e;
  let n = "",
    o = 0;
  for (; o < e.length;) {
    const a = e.slice(o),
      s = e[o],
      d =
        (o === 0 || !unquotedNameCharPattern.test(e[o - 1])) &&
        a.match(sheetReferencePrefixPattern);
    if (d) {
      ((n += d[0]), (o += d[0].length));
      continue;
    }
    if (s === '"' || s === "'") {
      const p = o++;
      for (; o < e.length;)
        if (e[o++] === s)
          if (e[o] === s) o++;
          else break;
      n += e.slice(p, o);
      continue;
    }
    if (s === "[") {
      const p = o++;
      let m = 1;
      for (; o < e.length && m;)
        (e[o] === "[" && m++, e[o] === "]" && m--, o++);
      n += e.slice(p, o);
      continue;
    }
    const u = o === 0 || !unquotedNameCharPattern.test(e[o - 1]),
      // A1:B2 is handled as one reference, so that a corner pushed off the
      // grid invalidates the whole range instead of half of it.
      x = u
        ? /^(\$?)([A-Za-z]{1,3})(\$?)([1-9]\d*):(\$?)([A-Za-z]{1,3})(\$?)([1-9]\d*)/.exec(
            a,
          )
        : null,
      l = u ? /^(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})/.exec(a) : null,
      c = u ? /^(\$?)([1-9]\d*):(\$?)([1-9]\d*)/.exec(a) : null,
      h = u ? /^(\$?)([A-Za-z]{1,3})(\$?)([1-9]\d*)/.exec(a) : null,
      f = x ?? l ?? c ?? h;
    if (
      f &&
      !unquotedNameCharPattern.test(a[f[0].length] ?? "") &&
      !/^\s*[(!]/.test(a.slice(f[0].length))
    ) {
      const joinRange = (...parts) =>
        parts.includes("#REF!")
          ? "#REF!"
          : parts.length === 4
            ? `${parts[0]}${parts[1]}:${parts[2]}${parts[3]}`
            : `${parts[0]}:${parts[1]}`;
      if (x)
        n += joinRange(
          shiftColumnReference(x[1], x[2], t),
          offsetRowReference(x[3], x[4], r),
          shiftColumnReference(x[5], x[6], t),
          offsetRowReference(x[7], x[8], r),
        );
      else if (l)
        n += joinRange(
          shiftColumnReference(l[1], l[2], t),
          shiftColumnReference(l[3], l[4], t),
        );
      else if (c)
        n += joinRange(
          offsetRowReference(c[1], c[2], r),
          offsetRowReference(c[3], c[4], r),
        );
      else if (h) {
        const p = shiftColumnReference(h[1], h[2], t),
          m = offsetRowReference(h[3], h[4], r);
        n += p === "#REF!" || m === "#REF!" ? "#REF!" : p + m;
      }
      o += f[0].length;
      continue;
    }
    ((n += s), o++);
  }
  return n;
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
