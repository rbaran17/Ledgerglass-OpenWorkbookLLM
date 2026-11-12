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
import { hr, iu } from "../01-core-schemas-and-hashing.js";
import { Mt } from "../02-financial-functions.js";
import { DQ, _Q, f6, p6 } from "./01-cell-styles-and-themes.js";

const jR = "ledgerglass-analysis-v1",
  ig = Object.freeze({
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
  th = number().finite(),
  Qv = th.min(-1e9).max(1e9),
  V4 = strictObject({
    sheetId: string().min(1).max(100),
    address: string().regex(/^[A-Z]{1,3}[1-9]\d{0,6}$/),
  }),
  wU = string()
    .trim()
    .min(1)
    .max(31)
    .refine(
      (e) => !/[\\/?*\[\]:\u0000-\u001f]/.test(e),
      "Invalid output worksheet name",
    ),
  yoe = strictObject({
    kind: literal("solver"),
    outputName: wU,
    objective: V4,
    direction: _enum(["maximize", "minimize", "target"]),
    target: Qv.optional(),
    variables: array(
      strictObject({
        cell: V4,
        lower: Qv,
        upper: Qv,
        domain: _enum(["continuous", "integer", "binary"]),
      }),
    )
      .min(1)
      .max(6),
    constraints: array(
      strictObject({
        left: V4,
        operator: _enum(["<=", ">=", "="]),
        right: union([Qv, V4]),
      }),
    ).max(16),
    tolerance: th.min(1e-10).max(1e-5).default(1e-8),
    maxCandidates: number().int().min(1).max(5e4).default(5e4),
  }),
  woe = strictObject({
    kind: literal("statistics"),
    outputName: wU,
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
    bins: array(Qv).min(1).max(50).optional(),
    distribution: _enum(["normal", "binomial", "poisson"]).optional(),
    mean: Qv.optional(),
    standardDeviation: th.positive().max(1e9).optional(),
    trials: number().int().min(0).max(1e3).optional(),
    probability: th.min(0).max(1).optional(),
    lambda: th.positive().max(1e3).optional(),
    forecast: _enum(["linear", "moving-average", "exponential"]).optional(),
    horizon: number().int().min(1).max(24).optional(),
    window: number().int().min(2).max(100).optional(),
    alpha: th.positive().max(1).optional(),
  }),
  xU = discriminatedUnion("kind", [yoe, woe]),
  xoe = union([string().max(32767), th, boolean(), _null()]),
  Eoe = strictObject({
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
    rows: array(array(xoe).max(16)).max(3e3),
    solution: array(th).max(6).optional(),
    objective: th.optional(),
    baseline: array(th).max(6).optional(),
  }),
  sk = string().regex(/^[a-f0-9]{64}$/),
  Coe = strictObject({
    engine: literal(jR),
    request: xU,
    result: Eoe,
    definitionHash: sk,
    sourceHash: sk,
    outputHash: sk,
    outputRange: string().max(50),
    createdAt: string().datetime(),
    sourceRevision: number().int().nonnegative(),
    includeHidden: boolean(),
    applied: boolean(),
  });
function EU(e) {
  const r = xU.safeParse(e);
  if (!r.success)
    throw Error(
      `Invalid analysis request: ${r.error.issues[0]?.path.join(".")}: ${r.error.issues[0]?.message}`,
    );
  const t = r.data;
  if (t.kind === "solver") {
    if (t.direction === "target" && t.target === undefined)
      throw Error("Enter a finite target value.");
    if (
      new Set(t.variables.map((a) => JSON.stringify(a.cell))).size !==
      t.variables.length
    )
      throw Error("Changing cells must be unique.");
    if (
      t.variables.some(
        (a) =>
          a.lower > a.upper ||
          (a.domain === "binary" && (a.lower < 0 || a.upper > 1)),
      )
    )
      throw Error(
        "Variable bounds must be ordered; binary bounds must lie within zero and one.",
      );
    if (
      t.variables.some((a) => a.domain === "continuous") &&
      t.variables.some((a) => a.domain !== "continuous")
    )
      throw Error(
        "Mixed integer and continuous models are outside this selected Solver subset.",
      );
  }
  return t;
}
function Aoe(e, r) {
  const t = e;
  let n = 0;
  for (const u of r.sheets) {
    const l = t.sheets.find((m) => m.id === u.id)?.analysisReport;
    if (l === undefined) continue;
    if (((n += JSON.stringify(l).length), n > ig.metadataCharacters))
      throw Error("Analysis metadata exceeds 2,000,000 characters.");
    const c = Coe.safeParse(l);
    if (!c.success)
      throw Error(`Invalid analysis report: ${c.error.issues[0]?.message}`);
    const h = c.data;
    if (
      ((h.request = EU(h.request)),
      h.definitionHash !== hr(h.request) ||
        !h.result.success ||
        !["complete", "optimal"].includes(h.result.status) ||
        h.result.rows.reduce((m, v) => m + v.length, 0) > ig.outputCells)
    )
      throw Error("Analysis report metadata is inconsistent.");
    const f = /^([A-Z]{1,3})([1-9]\d{0,6}):([A-Z]{1,3})([1-9]\d{0,6})$/.exec(
        h.outputRange,
      ),
      p = (m) => [...m].reduce((v, g) => v * 26 + g.charCodeAt(0) - 64, 0);
    if (
      !f ||
      p(f[3]) > 16384 ||
      Number(f[4]) > 1048576 ||
      p(f[3]) < p(f[1]) ||
      Number(f[4]) < Number(f[2]) ||
      (p(f[3]) - p(f[1]) + 1) * (Number(f[4]) - Number(f[2]) + 1) >
        ig.outputCells
    )
      throw Error("Invalid or oversized analysis output range.");
    if (
      (h.request.kind === "solver" &&
        (h.result.status !== "optimal" ||
          h.result.solution?.length !== h.request.variables.length ||
          h.result.baseline?.length !== h.request.variables.length ||
          h.result.objective === undefined)) ||
      (h.request.kind === "statistics" &&
        (h.result.status !== "complete" || h.applied))
    )
      throw Error("Analysis result does not match its request.");
    u.analysisReport = h;
  }
}
const CU = "ledgerglass-report-output-v1",
  Vv = Object.freeze({
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
  Tc = (e) => string().max(e),
  Ap = string().min(1).max(200),
  zE = string().regex(/^[a-f0-9]{64}$/),
  AU = string().datetime();
const W5 = (e) => {
  const { hash, ...t } = e;
  return hr(t);
};
const t0 = Object.freeze({
    expressionCharacters: 4096,
    syntaxNodes: 512,
    depth: 48,
    calculatedColumns: 16,
  }),
  H5 = Object.freeze({
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
class LR extends Error {
  constructor(r, t) {
    (super(t), (this.code = r), (this.name = "DaxError"));
  }
}
const on = (e, r) => {
  throw new LR(e, `DAX: ${r}`);
};
function UO(e) {
  if (typeof e != "string" || !e.trim() || e.length > t0.expressionCharacters)
    return on("DAX_LIMIT", "Enter an expression of 1–4,096 characters.");
  const r = e.trim().replace(/^=/, ""),
    t = [];
  let n = 0;
  const o = (m, v, g) => {
    const b = n++;
    let y = "",
      w = false;
    for (; n < r.length;) {
      const E = r[n++];
      if (E === v)
        if (r[n] === v) ((y += v), n++);
        else {
          w = true;
          break;
        }
      else y += E;
    }
    (w || on("DAX_SYNTAX", `Unclosed ${m} at position ${b + 1}.`),
      t.push({ kind: g, value: y, position: b }));
  };
  for (; n < r.length;) {
    if (/\s/.test(r[n])) {
      n++;
      continue;
    }
    if (r[n] === '"') {
      o('"', '"', "string");
      continue;
    }
    if (r[n] === "'") {
      o("'", "'", "quoted");
      continue;
    }
    if (r[n] === "[") {
      o("[", "]", "bracket");
      continue;
    }
    const m = r.slice(n),
      v = /^(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?/.exec(m),
      g = /^[A-Za-z_][A-Za-z0-9_]*/.exec(m),
      b = /^(?:==|<>|<=|>=|&&|\|\||[+\-*/^&=<>(),])/.exec(m),
      y = v ?? g ?? b;
    (y ||
      on(
        "DAX_UNSUPPORTED",
        `Unsupported syntax at position ${n + 1}. Use the documented DAX subset; worksheet and query syntax are not evaluated.`,
      ),
      t.push({
        kind: v ? "number" : g ? "identifier" : "op",
        value: y[0],
        position: n,
      }),
      (n += y[0].length),
      t.length > t0.syntaxNodes * 3 &&
        on("DAX_LIMIT", "Expression token budget exceeded."));
  }
  t.push({ kind: "end", value: "", position: n });
  let a = 0,
    s = 0;
  const d = () => t[a],
    u = () => t[a++],
    l = (m) => {
      u().value !== m && on("DAX_SYNTAX", `Expected ${m}.`);
    },
    c = (m) => (
      ++s > t0.syntaxNodes &&
        on("DAX_LIMIT", "Expression exceeds 512 syntax nodes."),
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
      v > t0.depth && on("DAX_LIMIT", "Expression nesting exceeds 48 levels.");
      const g = u();
      let b;
      if (g.kind === "number") {
        const y = Number(g.value);
        (Number.isFinite(y) ||
          on("DAX_NUMBER", "Numeric literal exceeds finite precision."),
          (b = c({ kind: "literal", value: y })));
      } else if (g.kind === "string")
        b = c({ kind: "literal", value: g.value });
      else if (g.value === "(") ((b = f(0, v + 1)), l(")"));
      else if (
        g.value === "+" ||
        g.value === "-" ||
        (g.kind === "identifier" && g.value.toUpperCase() === "NOT")
      )
        b = c({
          kind: "unary",
          op: g.value.toUpperCase(),
          value: f(g.value.toUpperCase() === "NOT" ? 3 : 8, v + 1),
        });
      else if (g.kind === "bracket") b = c({ kind: "measure", name: g.value });
      else if (g.kind === "identifier" || g.kind === "quoted")
        if (d().value === "(" && g.kind === "identifier") {
          const y = g.value.toUpperCase();
          (Object.prototype.hasOwnProperty.call(H5, y) ||
            on(
              "DAX_UNSUPPORTED",
              `Unsupported function ${g.value}. No worksheet fallback is used.`,
            ),
            u());
          const w = [];
          if (d().value !== ")")
            do {
              if ((w.push(f(0, v + 1)), d().value !== ",")) break;
              u();
            } while (true);
          l(")");
          const E = H5[y];
          ((w.length < E[0] || w.length > E[1]) &&
            on(
              "DAX_ARGUMENT",
              `${y} expects ${E[0] === E[1] ? E[0] : `${E[0]}–${E[1]}`} arguments in this subset.`,
            ),
            (b = c({ kind: "call", name: y, args: w })));
        } else
          d().kind === "bracket"
            ? (b = c({ kind: "field", table: g.value, column: u().value }))
            : g.kind === "identifier" &&
                ["TRUE", "FALSE"].includes(g.value.toUpperCase())
              ? (b = c({
                  kind: "literal",
                  value: g.value.toUpperCase() === "TRUE",
                }))
              : (b = c({ kind: "table", name: g.value }));
      else
        return on(
          "DAX_SYNTAX",
          `Expected an expression at position ${g.position + 1}.`,
        );
      for (
        ;
        d().kind === "op" && h[d().value] !== undefined && h[d().value] >= m;
      ) {
        const y = u().value,
          w = h[y];
        b = c({ kind: "binary", op: y, left: b, right: f(w + 1, v + 1) });
      }
      return b;
    },
    p = f();
  return (
    d().kind !== "end" &&
      on(
        "DAX_UNSUPPORTED",
        `Unexpected ${d().value}. VAR, RETURN, EVALUATE, worksheet references and unsupported syntax are not evaluated.`,
      ),
    p
  );
}
function Kh(e, r) {
  (r(e),
    e.kind === "call"
      ? e.args.forEach((t) => Kh(t, r))
      : e.kind === "binary"
        ? (Kh(e.left, r), Kh(e.right, r))
        : e.kind === "unary" && Kh(e.value, r));
}
const PR = "ledgerglass-columnar-v1",
  ta = Object.freeze({
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
  S1 = string()
    .min(1)
    .max(100)
    .regex(/^[A-Za-z0-9_.-]+$/),
  jg = string()
    .trim()
    .min(1)
    .max(128)
    .refine((e) => !/[\u0000-\u001f\u007f]/.test(e)),
  Ioe = strictObject({
    name: jg,
    type: _enum(["text", "number", "boolean", "date"]),
    nullable: boolean(),
    expression: string().trim().min(1).max(t0.expressionCharacters).optional(),
  }),
  _oe = strictObject({
    id: S1,
    name: jg,
    sheetId: string().min(1).max(100),
    range: string().min(1).max(50),
    tableId: string().min(1).max(100).optional(),
    columns: array(Ioe).min(1).max(ta.columns),
    dateColumn: jg.optional(),
  }),
  VE = strictObject({ tableId: S1, column: jg }),
  Doe = strictObject({
    id: S1,
    one: VE,
    many: VE,
    cardinality: _enum(["one-to-many", "one-to-one"]),
    direction: literal("single"),
  }),
  Ooe = strictObject({
    id: S1,
    name: jg,
    tableId: S1,
    column: jg.optional(),
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
    expression: string().trim().min(1).max(t0.expressionCharacters).optional(),
  }),
  SU = strictObject({
    engine: literal(PR),
    name: jg,
    tables: array(_oe).min(1).max(ta.tables),
    relationships: array(Doe).max(ta.relationships),
    measures: array(Ooe).max(ta.measures),
  }),
  Moe = union([string().max(32767), number().finite(), boolean(), _null()]),
  joe = strictObject({
    ...VE.shape,
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
    value: Moe.optional(),
  }),
  TU = strictObject({
    groupBy: array(VE).max(ta.groupFields),
    filters: array(joe).max(ta.filters),
    measureIds: array(S1).max(ta.measures),
  }),
  mp = (e, r) => {
    if (new Set(e.map((t) => t.toLowerCase())).size !== e.length)
      throw Error(`${r} must be unique, including case variants.`);
  };
function WE(e) {
  const r = SU.safeParse(e);
  if (!r.success)
    throw Error(
      `Invalid data model: ${r.error.issues[0]?.path.join(".")}: ${r.error.issues[0]?.message}`,
    );
  const t = r.data;
  if (
    (mp(
      t.tables.map((s) => s.id),
      "Model table IDs",
    ),
    mp(
      t.tables.map((s) => s.name),
      "Model table names",
    ),
    mp(
      t.relationships.map((s) => s.id),
      "Relationship IDs",
    ),
    mp(
      t.measures.map((s) => s.id),
      "Measure IDs",
    ),
    mp(
      t.measures.map((s) => s.name),
      "Measure names",
    ),
    t.tables.reduce(
      (s, d) => s + d.columns.filter((u) => u.expression !== undefined).length,
      0,
    ) > t0.calculatedColumns)
  )
    throw Error("The model supports at most 16 DAX calculated columns.");
  for (const s of t.tables) {
    if (
      (mp(
        s.columns.map((d) => d.name),
        `Columns in ${s.name}`,
      ),
      !s.columns.some((d) => d.expression === undefined))
    )
      throw Error(
        "A model table requires at least one physical source column.",
      );
    if (
      s.dateColumn &&
      !s.columns.some(
        (d) =>
          d.name === s.dateColumn &&
          d.type === "date" &&
          !d.nullable &&
          d.expression === undefined,
      )
    )
      throw Error(
        "A marked calendar requires a nonnullable physical date column.",
      );
  }
  const n = (s) => {
    const d = t.tables.find((l) => l.id === s.tableId),
      u = d?.columns.find((l) => l.name === s.column);
    if (!u) throw Error(`Unknown model field ${s.tableId}.${s.column}.`);
    return u;
  };
  for (const s of t.relationships) {
    if (s.one.tableId === s.many.tableId)
      throw Error(
        "Self relationships are unsupported; use distinct role-playing tables.",
      );
    if (n(s.one).type !== n(s.many).type)
      throw Error(`Relationship ${s.id} requires matching key types.`);
    if (n(s.one).expression !== undefined || n(s.many).expression !== undefined)
      throw Error(
        "Relationship keys must be physical source columns, not DAX calculated columns.",
      );
  }
  for (const s of t.measures) {
    if (!t.tables.some((d) => d.id === s.tableId))
      throw Error(`Unknown measure table ${s.tableId}.`);
    if (s.operation === "dax") {
      if (!s.expression || s.column !== undefined)
        throw Error(
          "A DAX measure requires an expression and no aggregate column.",
        );
    } else {
      if (s.expression !== undefined)
        throw Error(
          "Unrecognized expression for an aggregate measure; only DAX measures may contain it.",
        );
      if (s.operation === "countRows") {
        if (s.column !== undefined)
          throw Error("Count rows does not take a column.");
      } else {
        const d = n({ tableId: s.tableId, column: s.column ?? "" });
        if (
          ["sum", "average", "min", "max"].includes(s.operation) &&
          d.type !== "number"
        )
          throw Error(`Measure ${s.name} requires a number column.`);
      }
    }
  }
  return t;
}
function UR(e, r) {
  var a;
  const t = TU.safeParse(e);
  if (!t.success)
    throw Error(`Invalid model query: ${t.error.issues[0]?.message}`);
  const n = t.data;
  (mp(n.measureIds, "Selected measures"),
    mp(
      n.groupBy.map((s) => JSON.stringify(s)),
      "Group fields",
    ));
  for (const s of [...n.groupBy, ...n.filters])
    if (!(
      (a = r.tables.find((d) => d.id === s.tableId)) != null &&
      a.columns.some((d) => d.name === s.column)
    ))
      throw Error(`Unknown query field ${s.tableId}.${s.column}.`);
  if (n.measureIds.some((s) => !r.measures.some((d) => d.id === s)))
    throw Error("A selected model measure is missing.");
  return n;
}
const lk = string().regex(/^[a-f0-9]{64}$/),
  Foe = strictObject({
    definition: SU,
    query: TU,
    sourceHash: lk,
    definitionHash: lk,
    outputHash: lk,
    outputRange: string().min(1).max(50),
    createdAt: string().datetime(),
    sourceRevision: number().int().nonnegative(),
    includeHidden: boolean(),
    totals: array(number().finite().nullable()).max(ta.measures),
  });
function Boe(e, r) {
  const t = e;
  let n = 0;
  if (t.dataModel !== undefined) {
    if (((n += JSON.stringify(t.dataModel).length), n > ta.metadataCharacters))
      throw Error("Data model metadata exceeds 2,000,000 characters.");
    r.dataModel = WE(t.dataModel);
  }
  for (const s of r.sheets) {
    const d = t.sheets.find((c) => c.id === s.id)?.dataModelReport;
    if (d === undefined) continue;
    if (((n += JSON.stringify(d).length), n > ta.metadataCharacters))
      throw Error("Data model metadata exceeds 2,000,000 characters.");
    const u = Foe.safeParse(d);
    if (!u.success)
      throw Error(`Invalid model report: ${u.error.issues[0]?.message}`);
    const l = u.data;
    if (
      ((l.definition = WE(l.definition)),
      (l.query = UR(l.query, l.definition)),
      l.definitionHash !== hr(l.definition) ||
        l.totals.length !== l.query.measureIds.length)
    )
      throw Error(
        "Model report definition or measure metadata is inconsistent.",
      );
    s.dataModelReport = l;
  }
}
const Cd = Object.freeze({
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
  NU = ["json", "xml", "pdf", "png", "jpeg", "html"],
  Il = (e) => string().max(e),
  RU = string().regex(/^[a-f0-9]{64}$/),
  IU = union([Il(32767), number().finite(), boolean(), _null()]),
  Loe = object({
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
const Uoe = object({
    value: IU,
    type: _enum(["string", "number", "boolean", "null", "array", "object"]),
    path: Il(2048),
    lexical: Il(32767).optional(),
    declaredType: Il(200).optional(),
    missing: boolean().optional(),
    confidence: number().min(0).max(1),
    region: Loe.optional(),
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
  _U = object({
    id: Il(200).min(1),
    name: Il(200).min(1),
    path: Il(2048),
    ambiguous: boolean(),
    columns: array(object({ name: Il(200).min(1), path: Il(2048) }).strict())
      .min(1)
      .max(Cd.columns),
    rows: array(array(Uoe).max(Cd.columns)).min(1).max(Cd.rows),
  })
    .strict()
    .superRefine((e, r) => {
      (new Set(e.columns.map((t) => t.path)).size !== e.columns.length ||
        e.rows.some((t) => t.length !== e.columns.length)) &&
        r.addIssue({
          code: "custom",
          message: "Table columns and rows must align with unique paths.",
        });
    }),
  HE = object({
    fileName: Il(255)
      .min(1)
      .refine((e) => !/[\\/\x00-\x1f\x7f\u202a-\u202e\u2066-\u2069]/.test(e)),
    format: _enum(NU),
    byteLength: number().int().min(1).max(Cd.bytes),
    sha256: RU,
    base64: Il(Math.ceil(Cd.bytes / 3) * 4),
  }).strict(),
  DU = object({
    schemaVersion: literal(1),
    fileName: HE.shape.fileName,
    format: _enum(NU),
    byteLength: HE.shape.byteLength,
    sha256: RU,
    tables: array(_U).max(Cd.tables),
    method: _enum(["native-json", "native-xml", "document-adapter"]),
    extractor: Il(200).min(1),
    pageCount: number().int().min(1).max(1e4).optional(),
    warnings: array(Il(2048)).max(30),
  }).strict(),
  OU = object({
    decision: _enum(["accept", "exclude", "correct"]),
    value: IU.optional(),
  })
    .strict()
    .refine(
      (e) => e.decision !== "correct" || e.value !== undefined,
      "A corrected value is required.",
    ),
  $oe = object({
    id: Il(100).min(1),
    sheetId: Il(100).min(1),
    sheetName: Il(31).min(1),
    importedAt: string().datetime(),
    original: HE,
    method: DU.shape.method,
    extractor: Il(200).min(1),
    pageCount: number().int().min(1).max(1e4).optional(),
    table: _U,
    resolutions: record(string().regex(/^\d+:\d+$/), OU),
    ambiguityReviewed: boolean(),
  }).strict();
function kC(e) {
  const r = DU.parse(e);
  let t = 0;
  if (new Set(r.tables.map((n) => n.id)).size !== r.tables.length)
    throw Error("Source table IDs must be unique.");
  for (const n of r.tables)
    for (const o of n.rows)
      for (const a of o) {
        if (++t > Cd.cells) throw Error("Source cell budget exceeded.");
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
function $R(e) {
  const r = HE.parse(e),
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
  if (iu(o) !== r.sha256) throw Error("Retained source checksum mismatch.");
  return o;
}
function zoe(e, r, t) {
  let n = "";
  for (let a = 0; a < e.length; a += 8192)
    n += String.fromCharCode(...e.subarray(a, a + 8192));
  const o = {
    fileName: r,
    format: t,
    byteLength: e.length,
    sha256: iu(e),
    base64: btoa(n),
  };
  return ($R(o), o);
}
function zR(e, r) {
  if (e.confidence < Cd.confidence && !r)
    throw Error("Review every low-confidence cell before importing.");
  if (!r) return e.value;
  const t = OU.parse(r);
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
function Voe(e, r) {
  if (e.sourceImports === undefined) return;
  if (!Array.isArray(e.sourceImports) || e.sourceImports.length > Cd.imports)
    throw Error("Source import count exceeded.");
  if (
    e.sourceImports.some((o) => !o || typeof o != "object" || !uk.has(o)) &&
    new TextEncoder().encode(JSON.stringify(e.sourceImports)).length >
      24 * 1024 * 1024
  )
    throw Error("Source provenance metadata budget exceeded.");
  let t = 0;
  const n = new Set();
  if (
    ((r.sourceImports = e.sourceImports.map((o) => {
      if (o && typeof o == "object" && uk.has(o)) {
        const d = o;
        if (((t += d.original.byteLength), n.has(d.id)))
          throw Error("Duplicate source import ID.");
        return (n.add(d.id), d);
      }
      const a = $oe.parse(o);
      if (
        ($R(a.original),
        kC({
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
        (a.table.rows.forEach((d, u) =>
          d.forEach((l, c) => {
            const h = `${u}:${c}`;
            (s.add(h), zR(l, a.resolutions[h]));
          }),
        ),
        Object.keys(a.resolutions).some((d) => !s.has(d)))
      )
        throw Error("Unknown source review cell.");
      if (n.has(a.id)) throw Error("Duplicate source import ID.");
      return (n.add(a.id), (t += a.original.byteLength), MU(a), uk.add(a), a);
    })),
    t > Cd.retainedBytes)
  )
    throw Error("Retained source byte budget exceeded.");
}
const VR = Object.freeze({
    maxBytes: 10 * 1024 * 1024,
    maxParts: 2e3,
    maxDetails: 1e5,
  }),
  q5 = new WeakSet();
function SC(e) {
  if (
    !e ||
    typeof e != "object" ||
    e.schemaVersion !== 1 ||
    typeof e.fileName != "string" ||
    !/\.xlsx$/i.test(e.fileName) ||
    e.fileName.length > 255 ||
    /[\\/\x00-\x1f\x7f]/.test(e.fileName) ||
    !Number.isSafeInteger(e.byteLength) ||
    e.byteLength < 22 ||
    e.byteLength > VR.maxBytes ||
    typeof e.sha256 != "string" ||
    !/^[a-f0-9]{64}$/.test(e.sha256) ||
    typeof e.base64 != "string" ||
    e.base64.length !== Math.ceil(e.byteLength / 3) * 4 ||
    !/^[A-Za-z0-9+/]*={0,2}$/.test(e.base64)
  )
    throw Error(
      "Invalid retained XLSX source metadata or 10 MiB source limit exceeded.",
    );
  let r;
  try {
    r = atob(e.base64);
  } catch {
    throw Error("Invalid retained XLSX source encoding.");
  }
  if (btoa(r) !== e.base64 || r.length !== e.byteLength)
    throw Error("Retained XLSX source byte length or encoding changed.");
  const t = Uint8Array.from(r, (n) => n.charCodeAt(0));
  if (iu(t) !== e.sha256)
    throw Error("Retained XLSX source checksum does not match its bytes.");
  return t;
}
function Woe(e, r) {
  if (e.byteLength > VR.maxBytes)
    throw Error("Original XLSX exceeds the 10 MiB retention limit.");
  let t = "";
  for (let o = 0; o < e.length; o += 8192)
    t += String.fromCharCode(...e.subarray(o, o + 8192));
  const n = {
    schemaVersion: 1,
    fileName: r,
    byteLength: e.length,
    sha256: iu(e),
    base64: btoa(t),
  };
  return (SC(n), Object.freeze(n), q5.add(n), n);
}
function Hoe(e, r) {
  if (e.xlsxSource === undefined) return;
  const t = e.xlsxSource;
  if (t && typeof t == "object" && q5.has(t)) {
    r.xlsxSource = t;
    return;
  }
  (SC(t),
    (r.xlsxSource = Object.freeze({
      schemaVersion: 1,
      fileName: t.fileName,
      byteLength: t.byteLength,
      sha256: t.sha256,
      base64: t.base64,
    })),
    q5.add(r.xlsxSource));
}
function WR(e) {
  if (!e.xlsxSource)
    throw Error(
      "This workbook has no retained XLSX source. Open an XLSX file to retain its original bytes.",
    );
  return {
    bytes: SC(e.xlsxSource),
    fileName: e.xlsxSource.fileName,
    sha256: e.xlsxSource.sha256,
  };
}
const HR = ["xls", "xlsb", "xlsm", "xltx", "xltm", "ods"],
  $O = (e) =>
    typeof e == "string" &&
    e.length > 0 &&
    e.length <= 255 &&
    !/[\\/\x00-\x1f\x7f\u202a-\u202e\u2066-\u2069]/.test(e),
  jU = new WeakSet();
function qR(e) {
  if (
    !e ||
    typeof e != "object" ||
    !HR.includes(e.format) ||
    !$O(e.fileName) ||
    !e.fileName.toLowerCase().endsWith("." + e.format) ||
    !e.conversion ||
    !["template-adapter", "external-conversion"].includes(
      e.conversion.method,
    ) ||
    (e.conversion.method === "template-adapter" && e.format !== "xltx") ||
    !$O(e.conversion.convertedFileName) ||
    !/\.xlsx$/i.test(e.conversion.convertedFileName) ||
    typeof e.conversion.convertedSha256 != "string" ||
    !/^[a-f0-9]{64}$/.test(e.conversion.convertedSha256) ||
    typeof e.conversion.reviewedAt != "string" ||
    !/^\d{4}-\d\d-\d\dT\d\d:\d\d:\d\d\.\d{3}Z$/.test(e.conversion.reviewedAt) ||
    !Number.isFinite(Date.parse(e.conversion.reviewedAt)) ||
    e.conversion.equivalence !== "not-verified"
  )
    throw Error("Invalid alternate workbook source or conversion provenance.");
  return SC({ ...e, fileName: "retained.xlsx" });
}
function FU(e, r, t, n) {
  const o = {
    ...Woe(e, "retained.xlsx"),
    fileName: r,
    format: t,
    conversion: { ...n },
  };
  return (qR(o), Object.freeze(o.conversion), Object.freeze(o), jU.add(o), o);
}
function qoe(e, r) {
  if (e.alternateSource === undefined) return;
  const t = e.alternateSource;
  if (jU.has(t)) {
    r.alternateSource = t;
    return;
  }
  const n = qR(t);
  r.alternateSource = FU(n, t.fileName, t.format, t.conversion);
}
function Xoe(e) {
  const r = e.alternateSource;
  if (!r)
    throw Error("This workbook has no retained alternate-format original.");
  return {
    bytes: qR(r),
    fileName: r.fileName,
    sha256: r.sha256,
    format: r.format,
  };
}
const ol = Object.freeze({
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
  Goe = [
    "vba",
    "office-script",
    "office-script-container",
    "excel-container",
    "legacy-container",
  ];
function Zoe(e) {
  let r = "";
  for (let t = 0; t < e.length; t += 8192)
    r += String.fromCharCode(...e.subarray(t, t + 8192));
  return btoa(r);
}
function X5(e) {
  if (
    typeof e != "string" ||
    e.length > Math.ceil(ol.maxArtifactBytes / 3) * 4 ||
    !/^(?:[A-Za-z0-9+/]{4})*(?:[A-Za-z0-9+/]{2}==|[A-Za-z0-9+/]{3}=)?$/.test(e)
  )
    throw Error("Invalid automation artifact encoding or size.");
  const r = atob(e);
  if (btoa(r) !== e)
    throw Error("Automation artifact is not canonical base64.");
  return Uint8Array.from(r, (t) => t.charCodeAt(0));
}
const qh = (e, r, t) => {
    if (typeof e != "string" || !e.trim() || e.length > r)
      throw Error(`Invalid automation ${t}.`);
    return e;
  },
  ck = (e) => {
    if (typeof e != "string" || !/^[a-f0-9]{64}$/.test(e))
      throw Error("Invalid automation source hash.");
    return e;
  },
  Gb = (e, r, t) => {
    if (!Number.isSafeInteger(e) || Number(e) < r || Number(e) > t)
      throw Error("Invalid automation evidence count.");
    return Number(e);
  },
  zO = (e) => {
    const r = qh(e, 40, "timestamp");
    if (!/^\d{4}-\d\d-\d\dT/.test(r) || !Number.isFinite(Date.parse(r)))
      throw Error("Invalid automation timestamp.");
    return r;
  },
  Koe = (e) => {
    const r = qh(e, 40, "approved range"),
      t = /^([A-Z]{1,3})([1-9]\d*):([A-Z]{1,3})([1-9]\d*)$/.exec(r);
    if (!t) throw Error("Invalid automation approved range.");
    const n = (u) => [...u].reduce((l, c) => l * 26 + c.charCodeAt(0) - 64, 0),
      [o, a, s, d] = [n(t[1]), Number(t[2]), n(t[3]), Number(t[4])];
    if (
      o > s ||
      a > d ||
      s > 16384 ||
      d > 1048576 ||
      (s - o + 1) * (d - a + 1) > ol.maxCells
    )
      throw Error("Invalid automation approved range.");
    return r;
  };
function VO(e) {
  const r = qh(e, 255, "file name");
  if (/[\\/\x00-\x1f\x7f]/.test(r) || r === "." || r === "..")
    throw Error("Invalid automation file name.");
  return r;
}
function Yoe(e, r) {
  if (e.automationArtifacts !== undefined) {
    if (
      !Array.isArray(e.automationArtifacts) ||
      e.automationArtifacts.length > ol.maxArtifacts
    )
      throw Error("Keep at most 20 automation artifacts.");
    let t = 0;
    const n = new Set();
    r.automationArtifacts = e.automationArtifacts.map((o) => {
      if (!o || typeof o != "object" || Array.isArray(o))
        throw Error("Invalid automation artifact.");
      const a = o,
        s = qh(a.id, 100, "artifact ID");
      if (n.has(s)) throw Error("Duplicate automation artifact ID.");
      n.add(s);
      const d = X5(a.base64);
      if (
        ((t += d.length),
        !d.length ||
          d.length !== a.byteLength ||
          d.length > ol.maxArtifactBytes ||
          t > ol.maxStoredBytes)
      )
        throw Error(
          "Automation artifacts exceed their 2 MB per-file or 10 MB retained-byte budget.",
        );
      if (iu(d) !== ck(a.sha256))
        throw Error(
          "Automation artifact checksum does not match its preserved bytes.",
        );
      if (!Goe.includes(a.kind))
        throw Error("Unknown automation artifact kind.");
      return {
        id: s,
        fileName: VO(a.fileName),
        kind: a.kind,
        sha256: a.sha256,
        byteLength: d.length,
        base64: a.base64,
        retainedAt: zO(a.retainedAt),
        retainedBy: qh(a.retainedBy, 100, "retainer"),
      };
    });
  }
  if (e.automationRuns !== undefined) {
    if (
      !Array.isArray(e.automationRuns) ||
      e.automationRuns.length > ol.maxRuns
    )
      throw Error("Automation run evidence is limited to 100 runs.");
    const t = new Set();
    r.automationRuns = e.automationRuns.map((n) => {
      if (!n || typeof n != "object" || Array.isArray(n))
        throw Error("Invalid automation run evidence.");
      const o = n,
        a = qh(o.id, 100, "run ID");
      if (t.has(a)) throw Error("Duplicate automation run ID.");
      if (
        (t.add(a),
        !["vba", "office-script"].includes(o.language) ||
          o.adapter !== "literal-ranges-v1")
      )
        throw Error("Unknown automation migration adapter.");
      if (
        !Array.isArray(o.sourceLines) ||
        o.sourceLines.length > ol.maxCommands
      )
        throw Error("Invalid automation source lines.");
      const s = Gb(o.beforeRevision, 0, Number.MAX_SAFE_INTEGER - 2),
        d = Gb(o.afterRevision, 1, Number.MAX_SAFE_INTEGER - 1);
      if (d !== s + 1)
        throw Error("Automation run revisions are inconsistent.");
      if (o.sourceLines.length !== o.commands)
        throw Error(
          "Automation source-line evidence must match the command count.",
        );
      return {
        id: a,
        eventId: qh(o.eventId, 100, "event ID"),
        sourceName: VO(o.sourceName),
        sourceSha256: ck(o.sourceSha256),
        codeSha256: ck(o.codeSha256),
        language: o.language,
        adapter: o.adapter,
        actor: qh(o.actor, 100, "actor"),
        timestamp: zO(o.timestamp),
        beforeRevision: s,
        afterRevision: d,
        sheetId: qh(o.sheetId, 100, "sheet ID"),
        range: Koe(o.range),
        commands: Gb(o.commands, 1, ol.maxCommands),
        changedCells: Gb(o.changedCells, 0, ol.maxCells),
        sourceLines: o.sourceLines.map((u) => Gb(u, 1, ol.maxSourceBytes + 1)),
      };
    });
  }
}
function Joe(e, r) {
  if (e === undefined) return;
  if (!e || typeof e != "object" || Array.isArray(e))
    throw Error("Invalid delimited source footprint.");
  const t = e;
  if (
    !Number.isInteger(t.records) ||
    Number(t.records) < 1 ||
    Number(t.records) > Mt.maxRows ||
    !Number.isInteger(t.columns) ||
    Number(t.columns) < 1 ||
    Number(t.columns) > Mt.maxCols ||
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
  Qoe = [
    "deleted-cells",
    "deleted-sheet",
    "overwritten",
    "missing-sheet",
    "out-of-bounds",
  ],
  WO = (e) =>
    !!e &&
    typeof e == "object" &&
    !Array.isArray(e) &&
    [Object.prototype, null].includes(Object.getPrototypeOf(e)),
  HO = (e, r) => Object.keys(e).every((t) => r.includes(t)),
  dk = (e, r) =>
    typeof e == "string" &&
    !!e.trim() &&
    e.length <= r &&
    !/[\u0000-\u001f]/.test(e);
const Ls = {
    inputs: 32,
    outputs: 16,
    scenarios: 100,
    axisValues: 100,
    trials: 400,
    dependencyCells: 500,
    workUnits: 5e4,
    savedAnalyses: 40,
  },
  Vu = (e) => JSON.stringify([e.sheetId, e.address]),
  qE = (e) => typeof e == "number" && Number.isFinite(e);
function Yh(e, r) {
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
function Ep(e, r = 100) {
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
function l1(e, r, t = 1) {
  if (!Array.isArray(e) || e.length < t || e.length > r)
    throw Error(`What-if lists require ${t}–${r} entries.`);
  return Array.from(e);
}
function e1(e, r) {
  return l1(e, r).map((t) => {
    if (!qE(t))
      throw Error("Every what-if input value must be a finite number.");
    return t;
  });
}
function PU(e) {
  return e.kind === "scenarios"
    ? e.scenarios.map((r) => ({ name: r.name, inputs: r.values }))
    : e.kind === "one-variable"
      ? e.values.map((r) => ({ name: String(r), inputs: [r] }))
      : e.columnValues.flatMap((r) =>
          e.rowValues.map((t) => ({ name: `${t} / ${r}`, inputs: [t, r] })),
        );
}
const wg = {
    areaCells: 5e3,
    pages: 100,
    titleRows: 100,
    titleColumns: 50,
    breaksPerAxis: 100,
    renderedCells: 1e5,
  },
  XE = {
    header: `&[Workbook]
Engagement: &[Engagement] · Period: &[Period]`,
    footer: "Revision &[Revision] · &A · Page &P of &N",
  },
  GR = /&&|&\[(Workbook|Revision|Engagement|Period)\]|&[PNA]/g;
function UU(e, r) {
  return e === "Workbook"
    ? (r?.name ?? "[Workbook not available]")
    : e === "Revision"
      ? r
        ? String(r.revision)
        : "[Revision not available]"
      : e === "Engagement"
        ? r?.metadata?.engagement || "[Engagement not set]"
        : r?.metadata?.period || "[Period not set]";
}
function rae(e, r, t, n, o) {
  return e.replace(GR, (a, s) =>
    s
      ? UU(s, r)
      : a === "&&"
        ? "&"
        : a === "&A"
          ? t.name
          : String(a === "&P" ? n : o),
  );
}
function XO(e, r) {
  const t = e.replace(GR, (n, o) => {
    if (!o) return n;
    if (!r)
      throw Error(
        "Exporting print identity fields requires the current workbook snapshot.",
      );
    return UU(o, r).replace(/&/g, "&&");
  });
  if (t.length > 255)
    throw Error(
      "Resolved print identifiers exceed Excel's 255-character header/footer limit. Shorten the header, footer or workbook properties before exporting.",
    );
  return t;
}
function GE(e, r) {
  if (e !== undefined) {
    if (
      !Array.isArray(e) ||
      e.length > wg.breaksPerAxis ||
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
function nae(e) {
  return Object.freeze({
    workbookId: e.id,
    revision: e.revision,
    snapshotGuard: JSON.stringify(e),
  });
}
function $U(e, r) {
  if (
    r.workbookId !== e.id ||
    r.revision !== e.revision ||
    r.snapshotGuard !== JSON.stringify(e)
  )
    throw Error(
      "The workbook changed after page setup opened. Close and reopen page setup.",
    );
}
function KR(e) {
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
const VU = [
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
  sae = (e) =>
    !!e &&
    typeof e == "object" &&
    !Array.isArray(e) &&
    [Object.prototype, null].includes(Object.getPrototypeOf(e)),
  gp = () => {
    throw Error(
      "Invalid or over-capacity text cleanup evidence. Keep the original native backup.",
    );
  };
function L2(e, r) {
  (!sae(e) || Object.keys(e).some((t) => !r.includes(t))) && gp();
}
function WU(e) {
  return (
    e === null ||
    typeof e == "boolean" ||
    (typeof e == "number" && Number.isFinite(e)) ||
    (typeof e == "string" && e.length <= 32767)
  );
}
function jm(e, r) {
  return typeof e == "string" && e.length > 0 && e.length <= r;
}
function YR(e) {
  if (!Array.isArray(e) || !e.length || e.length > 12)
    throw Error("Choose 1–12 text cleanup steps.");
  return e.map(
    (r) => (
      L2(r, ["function", "args"]),
      (!VU.includes(r.function) ||
        !Array.isArray(r.args) ||
        r.args.length > 5 ||
        r.args.some((t) => t === null || !WU(t)) ||
        JSON.stringify(r.args).length > 4096) &&
        gp(),
      { function: r.function, args: [...r.args] }
    ),
  );
}
function lae(e, r) {
  var n;
  if (!((n = e.textCleanups) != null && n.length)) return r;
  const t = new Map(e.textCleanups.map((o) => [o.id, o]));
  for (const o of r.textCleanups ?? []) {
    const a = t.get(o.id);
    if (a && JSON.stringify(a) !== JSON.stringify(o))
      throw Error("Saved original text cleanup evidence cannot be changed.");
  }
  return {
    ...r,
    textCleanups: [
      ...e.textCleanups,
      ...(r.textCleanups ?? []).filter((o) => !t.has(o.id)),
    ],
  };
}
const Z5 = {
    inferencePatterns: 8e3,
    inferenceCandidates: 30,
    evidenceCharacters: 5e6,
  },
  ZE = number().int().min(0).max(49),
  Bh = array(ZE)
    .min(1)
    .max(50)
    .refine(
      (e) => new Set(e).size === e.length,
      "Choose distinct column positions.",
    ),
  qU = string().max(32767),
  JR = union([qU, number().finite(), boolean(), _null()]),
  XU = _enum(["none", "upper", "lower", "proper"]),
  uae = strictObject({
    column: ZE,
    mode: _enum(["whole", "token", "slice", "before", "after"]),
    delimiter: string().min(1).max(16).default(" "),
    index: number().int().min(-50).max(49).default(0),
    start: number().int().min(0).max(32767).default(0),
    length: number().int().min(1).max(32767).default(1),
  }),
  GU = strictObject({
    parts: array(uae).min(1).max(8),
    separator: string().max(100),
    prefix: string().max(1e3),
    suffix: string().max(1e3),
    casing: XU,
  }),
  ZU = strictObject({ row: number().int().min(1).max(1048576), value: qU }),
  cae = array(ZU)
    .min(2)
    .max(10)
    .refine(
      (e) => new Set(e.map((r) => r.row)).size === e.length,
      "Use examples from distinct source rows.",
    ),
  dae = discriminatedUnion("kind", [
    strictObject({
      kind: literal("trim"),
      columns: Bh,
      mode: _enum(["edges", "collapse"]),
    }),
    strictObject({ kind: literal("clean"), columns: Bh }),
    strictObject({ kind: literal("case"), columns: Bh, mode: XU }),
    strictObject({
      kind: literal("replace"),
      columns: Bh,
      find: string().min(1).max(1e3),
      replacement: string().max(1e3),
    }),
    strictObject({
      kind: literal("split"),
      column: ZE,
      mode: _enum(["delimited", "fixed"]),
      delimiters: string().max(16),
      collapse: boolean(),
      breaks: array(number().int().min(1).max(32766)).max(49),
    }),
    strictObject({
      kind: literal("concat"),
      columns: Bh,
      separator: string().max(100),
      name: string().min(1).max(100),
      skipBlank: boolean(),
    }),
    strictObject({
      kind: literal("convert"),
      columns: Bh,
      to: _enum(["text", "number", "integer", "boolean", "date"]),
      decimal: _enum([".", ","]),
      group: _enum(["", ".", ",", " ", "'"]),
      dateOrder: _enum(["MDY", "DMY"]),
    }),
    strictObject({
      kind: literal("fill"),
      columns: Bh,
      direction: _enum(["down", "up", "value"]),
      value: JR,
      groupColumns: array(ZE).max(50),
    }),
    strictObject({ kind: literal("select"), columns: Bh }),
    strictObject({
      kind: literal("pattern"),
      name: string().min(1).max(100),
      pattern: GU,
      examples: cae,
      inferenceColumns: Bh.max(8),
    }),
  ]),
  fae = strictObject({
    sourceSheetId: string().min(1).max(100),
    sourceRange: string().min(1).max(50),
    header: boolean(),
    outputName: string().min(1).max(31),
    steps: array(dae).max(12),
  });
function KU(e) {
  const r = fae.safeParse(e);
  if (!r.success)
    throw Error(
      `Invalid field transformation: ${r.error.issues[0]?.message ?? "review the recipe"}`,
    );
  const t = r.data;
  if (
    ((t.outputName = t.outputName.trim()),
    !t.outputName ||
      /[\u0000-\u001f\u007f\\/\[\]*?:]/.test(t.outputName) ||
      /^'|'$/.test(t.outputName))
  )
    throw Error("Choose a valid output worksheet name of 1–31 characters.");
  for (const o of t.steps) {
    if (o.kind === "split") {
      if (
        o.mode === "delimited" &&
        (![...o.delimiters].length ||
          [...o.delimiters].length > 8 ||
          /["\r\n\0]/.test(o.delimiters))
      )
        throw Error(
          "Choose 1–8 delimiter characters without quotes, line breaks or NUL.",
        );
      if (
        o.mode === "fixed" &&
        (!o.breaks.length ||
          o.breaks.some((a, s) => s > 0 && a <= o.breaks[s - 1]))
      )
        throw Error(
          "Fixed-width breaks must be distinct increasing character positions.",
        );
    }
    if (o.kind === "convert" && o.to !== "text" && o.decimal === o.group)
      throw Error("Decimal and grouping separators must differ.");
    if (
      o.kind === "fill" &&
      new Set(o.groupColumns).size !== o.groupColumns.length
    )
      throw Error("Choose distinct fill group columns.");
    if (o.kind === "fill" && o.columns.some((a) => o.groupColumns.includes(a)))
      throw Error(
        "Fill targets must differ from the group columns. Use a separate preceding step to prepare group keys.",
      );
  }
  return t;
}
function bae(e, r) {
  var n;
  if (!((n = e.fieldTransforms) != null && n.length)) return r;
  const t = new Map(e.fieldTransforms.map((o) => [o.id, o]));
  for (const o of r.fieldTransforms ?? [])
    if (t.has(o.id) && JSON.stringify(t.get(o.id)) !== JSON.stringify(o))
      throw Error(
        "Saved original field transformation evidence cannot be changed.",
      );
  return {
    ...r,
    fieldTransforms: [
      ...e.fieldTransforms,
      ...(r.fieldTransforms ?? []).filter((o) => !t.has(o.id)),
    ],
  };
}
const JU = new WeakSet(),
  NC = new WeakSet(),
  YE = new WeakMap(),
  QU = new WeakMap(),
  yae = new Set(["general", "number", "currency", "percent", "date"]),
  wae = (e, r) => Object.prototype.hasOwnProperty.call(e, r),
  Vh = (e) => !!e && typeof e == "object" && !Array.isArray(e),
  Hn = (e) => {
    throw new Error(e);
  },
  Bu = (e, r, t = 32767) =>
    typeof e == "string" && e.length <= t
      ? e
      : Hn(`${r} must be text of at most ${t} characters.`),
  P2 = (e, r, t, n) =>
    typeof e == "number" && Number.isInteger(e) && e >= r && e <= t
      ? e
      : Hn(`${n} must be an integer between ${r} and ${t}.`);
function ut(e) {
  P2(e, 0, Mt.maxCols - 1, "Column");
  let r = "";
  for (let t = e + 1; t > 0; t = Math.floor((t - 1) / 26))
    r = String.fromCharCode(65 + ((t - 1) % 26)) + r;
  return r;
}
function h6(e, r, t) {
  const n = DQ(r) + (e ? 0 : t);
  return n < 0 || n >= Mt.maxCols ? "#REF!" : e + ut(n);
}
function xd(e, r, t) {
  if (!Number.isSafeInteger(r) || !Number.isSafeInteger(t))
    throw new Error("Formula offsets must be integers.");
  if (!e.startsWith("=")) return e;
  let n = "",
    o = 0;
  for (; o < e.length;) {
    const a = e.slice(o),
      s = e[o],
      d = (o === 0 || !f6.test(e[o - 1])) && a.match(_Q);
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
    const u = o === 0 || !f6.test(e[o - 1]),
      l = u ? /^(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})/.exec(a) : null,
      c = u ? /^(\$?)([1-9]\d*):(\$?)([1-9]\d*)/.exec(a) : null,
      h = u ? /^(\$?)([A-Za-z]{1,3})(\$?)([1-9]\d*)/.exec(a) : null,
      f = l ?? c ?? h;
    if (
      f &&
      !f6.test(a[f[0].length] ?? "") &&
      !/^\s*[(!]/.test(a.slice(f[0].length))
    ) {
      if (l) n += `${h6(l[1], l[2], t)}:${h6(l[3], l[4], t)}`;
      else if (c) n += `${p6(c[1], c[2], r)}:${p6(c[3], c[4], r)}`;
      else if (h) {
        const p = h6(h[1], h[2], t),
          m = p6(h[3], h[4], r);
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
  $R,
  $U,
  AU,
  Aoe,
  Ap,
  B2,
  Boe,
  Bu,
  CU,
  Cd,
  EU,
  Ep,
  FU,
  GE,
  GR,
  GU,
  H5,
  HO,
  HR,
  Hn,
  Hoe,
  JR,
  JU,
  Joe,
  KR,
  KU,
  Kh,
  L2,
  LR,
  Ls,
  NC,
  P2,
  PR,
  PU,
  QU,
  Qoe,
  Tc,
  UO,
  UR,
  VR,
  VU,
  Vh,
  Voe,
  Vu,
  Vv,
  W5,
  WE,
  WO,
  WR,
  WU,
  X5,
  XE,
  XO,
  Xoe,
  YE,
  YR,
  Yh,
  Yoe,
  Z5,
  ZU,
  Zoe,
  bae,
  dk,
  e1,
  gp,
  ig,
  jR,
  jm,
  kC,
  l1,
  lae,
  nae,
  ol,
  on,
  qE,
  qoe,
  rae,
  t0,
  ta,
  ut,
  wae,
  wg,
  xd,
  yae,
  zE,
  zR,
  zoe,
};
