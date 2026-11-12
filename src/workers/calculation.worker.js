import * as E from "zod";
import __lib_fast_formula_parser from "fast-formula-parser";
var XR = Object.create;
var Qf = Object.defineProperty;
var JR = Object.getOwnPropertyDescriptor;
var YR = Object.getOwnPropertyNames;
var QR = Object.getPrototypeOf,
  eT = Object.prototype.hasOwnProperty;
var ye = (e, t) => () => (t || e((t = { exports: {} }).exports, t), t.exports),
  Tn = (e, t) => {
    for (var r in t) Qf(e, r, { get: t[r], enumerable: true });
  },
  tT = (e, t, r, o) => {
    if ((t && typeof t == "object") || typeof t == "function")
      for (let n of YR(t))
        !eT.call(e, n) &&
          n !== r &&
          Qf(e, n, {
            get: () => t[n],
            enumerable: !(o = JR(t, n)) || o.enumerable,
          });
    return e;
  };
var rT = (e, t, r) => (
  (r = e != null ? XR(QR(e)) : {}),
  tT(
    t || !e || !e.__esModule
      ? Qf(r, "default", { value: e, enumerable: true })
      : r,
    e,
  )
);

var qN = () => __lib_fast_formula_parser;
var ed = new WeakSet();
function dn(e) {
  return !!e && typeof e == "object" && ed.has(e);
}
function oi(e) {
  if (!e || typeof e != "object" || ed.has(e)) return e;
  for (let t of Object.values(e)) oi(t);
  return (Object.freeze(e), ed.add(e), e);
}
var KD = new TextEncoder();

function Fo(e) {
  if (
    !ArrayBuffer.isView(e) ||
    Object.prototype.toString.call(e) !== "[object Uint8Array]" ||
    e.length > 50 * 1024 * 1024
  )
    throw Error("Hashing requires at most 50 MB of bytes.");
  let t = [
      1116352408, 1899447441, 3049323471, 3921009573, 961987163, 1508970993,
      2453635748, 2870763221, 3624381080, 310598401, 607225278, 1426881987,
      1925078388, 2162078206, 2614888103, 3248222580, 3835390401, 4022224774,
      264347078, 604807628, 770255983, 1249150122, 1555081692, 1996064986,
      2554220882, 2821834349, 2952996808, 3210313671, 3336571891, 3584528711,
      113926993, 338241895, 666307205, 773529912, 1294757372, 1396182291,
      1695183700, 1986661051, 2177026350, 2456956037, 2730485921, 2820302411,
      3259730800, 3345764771, 3516065817, 3600352804, 4094571909, 275423344,
      430227734, 506948616, 659060556, 883997877, 958139571, 1322822218,
      1537002063, 1747873779, 1955562222, 2024104815, 2227730452, 2361852424,
      2428436474, 2756734187, 3204031479, 3329325298,
    ],
    r = [
      1779033703, 3144134277, 1013904242, 2773480762, 1359893119, 2600822924,
      528734635, 1541459225,
    ],
    o = new Uint8Array(Math.ceil((e.length + 9) / 64) * 64);
  (o.set(e), (o[e.length] = 128));
  let n = new DataView(o.buffer);
  n.setUint32(o.length - 4, e.length * 8);
  let i = new Uint32Array(64),
    a = (l, s) => (l >>> s) | (l << (32 - s));
  for (let l = 0; l < o.length; l += 64) {
    for (let v = 0; v < 16; v++) i[v] = n.getUint32(l + v * 4);
    for (let v = 16; v < 64; v++) {
      let g = i[v - 15],
        y = i[v - 2];
      i[v] =
        (i[v - 16] +
          (a(g, 7) ^ a(g, 18) ^ (g >>> 3)) +
          i[v - 7] +
          (a(y, 17) ^ a(y, 19) ^ (y >>> 10))) >>>
        0;
    }
    let [s, c, u, f, d, p, m, h] = r;
    for (let v = 0; v < 64; v++) {
      let g =
          (h +
            (a(d, 6) ^ a(d, 11) ^ a(d, 25)) +
            ((d & p) ^ (~d & m)) +
            t[v] +
            i[v]) >>>
          0,
        y =
          ((a(s, 2) ^ a(s, 13) ^ a(s, 22)) + ((s & c) ^ (s & u) ^ (c & u))) >>>
          0;
      ((h = m),
        (m = p),
        (p = d),
        (d = (f + g) >>> 0),
        (f = u),
        (u = c),
        (c = s),
        (s = (g + y) >>> 0));
    }
    [s, c, u, f, d, p, m, h].forEach((v, g) => (r[g] = (r[g] + v) >>> 0));
  }
  return r.map((l) => l.toString(16).padStart(8, "0")).join("");
}
var xs = Object.freeze({
    sources: 8,
    steps: 20,
    rows: 1e3,
    columns: 50,
    cells: 2e4,
    characters: 1e6,
    lineageLinks: 2e4,
    recipes: 20,
    versions: 20,
    metadataCharacters: 2e6,
    runMetadataCharacters: 8e6,
  }),
  XE = "ledgerglass-recipes-v1",
  Dt = E.string()
    .trim()
    .min(1)
    .max(128)
    .refine((e) => !/[\u0000-\u001f\u007f]/.test(e)),
  zo = E.string().min(1).max(100),
  A0 = E.union([
    E.string().max(32767),
    E.number().finite(),
    E.boolean(),
    E.null(),
  ]),
  ji = E.array(Dt)
    .max(50)
    .refine((e) => new Set(e).size === e.length, "Fields must be distinct."),
  N0 = E.strictObject({
    sheetId: zo,
    range: E.string().min(1).max(50),
    tableId: zo.optional(),
  }),
  HE = E.strictObject({
    field: Dt,
    operation: E.enum([
      "sum",
      "count",
      "average",
      "min",
      "max",
      "distinctCount",
    ]),
    name: Dt,
  }),
  j$ = E.strictObject({
    field: Dt,
    operator: E.enum([
      "equals",
      "notEquals",
      "contains",
      "startsWith",
      "endsWith",
      "gt",
      "gte",
      "lt",
      "lte",
      "blank",
      "notBlank",
    ]),
    value: A0.optional(),
  }),
  Br = { id: zo },
  Z$ = E.discriminatedUnion("kind", [
    E.strictObject({ ...Br, kind: E.literal("select"), fields: ji.min(1) }),
    E.strictObject({
      ...Br,
      kind: E.literal("rename"),
      mappings: E.array(E.strictObject({ from: Dt, to: Dt }))
        .min(1)
        .max(50),
    }),
    E.strictObject({
      ...Br,
      kind: E.literal("filter"),
      conditions: E.array(j$).min(1).max(20),
      join: E.enum(["and", "or"]),
    }),
    E.strictObject({
      ...Br,
      kind: E.literal("types"),
      fields: ji.min(1),
      to: E.enum(["text", "number", "integer", "boolean", "date"]),
      decimal: E.enum([".", ","]),
      group: E.enum(["", ".", ",", " ", "'"]),
      dateOrder: E.enum(["MDY", "DMY"]),
    }),
    E.strictObject({
      ...Br,
      kind: E.literal("replace"),
      fields: ji.min(1),
      mode: E.enum(["value", "text"]),
      find: A0,
      replacement: A0,
    }),
    E.strictObject({
      ...Br,
      kind: E.literal("derive"),
      name: Dt,
      expression: E.string().min(1).max(1e3),
    }),
    E.strictObject({
      ...Br,
      kind: E.literal("sort"),
      fields: E.array(
        E.strictObject({
          field: Dt,
          direction: E.enum(["ascending", "descending"]),
        }),
      )
        .min(1)
        .max(8),
    }),
    E.strictObject({
      ...Br,
      kind: E.literal("unpivot"),
      fields: ji.min(1),
      attribute: Dt,
      value: Dt,
      keepNulls: E.boolean(),
    }),
    E.strictObject({
      ...Br,
      kind: E.literal("group"),
      keys: ji,
      values: E.array(HE).min(1).max(20),
    }),
    E.strictObject({
      ...Br,
      kind: E.literal("pivot"),
      keys: ji,
      column: Dt,
      value: Dt,
      operation: HE.shape.operation,
    }),
    E.strictObject({ ...Br, kind: E.literal("append"), source: N0 }),
    E.strictObject({
      ...Br,
      kind: E.literal("merge"),
      source: N0,
      mode: E.enum(["inner", "left", "right", "full", "leftAnti", "rightAnti"]),
      keys: E.array(E.strictObject({ left: Dt, right: Dt }))
        .min(1)
        .max(8),
      prefix: Dt,
      normalization: E.strictObject({
        trim: E.boolean().optional(),
        collapseWhitespace: E.boolean().optional(),
        ignoreCase: E.boolean().optional(),
        unicode: E.enum(["NFC", "NFKC"]).optional(),
      }).optional(),
      fuzzy: E.strictObject({
        key: E.number().int().min(0).max(7),
        threshold: E.number().finite().min(0).max(1),
      }).optional(),
    }),
  ]),
  JE = E.strictObject({
    engine: E.literal(XE),
    source: N0,
    steps: E.array(Z$).max(xs.steps),
  }).superRefine((e, t) => {
    (new Set(e.steps.map((r) => r.id)).size !== e.steps.length &&
      t.addIssue({ code: "custom", message: "Step IDs must be unique." }),
      new Set(
        [
          e.source,
          ...e.steps.flatMap((r) =>
            r.kind === "append" || r.kind === "merge" ? [r.source] : [],
          ),
        ].map((r) => JSON.stringify(r)),
      ).size > xs.sources &&
        t.addIssue({
          code: "custom",
          message: "Use no more than eight recipe sources.",
        }));
  });
function G$(e) {
  return JSON.stringify(e, (t, r) =>
    r && typeof r == "object" && !Array.isArray(r)
      ? Object.fromEntries(
          Object.keys(r)
            .sort()
            .map((o) => [o, r[o]]),
        )
      : r,
  );
}
function Mn(e) {
  let t = G$(e);
  if (t.length > 16e6)
    throw Error(
      "Recipe fingerprint exceeds the 16-million-character evidence budget.",
    );
  return Fo(new Uint8Array(new TextEncoder().encode(t)));
}
var Zi = E.string().regex(/^[0-9a-f]{64}$/),
  Is = E.number()
    .int()
    .nonnegative()
    .max(Number.MAX_SAFE_INTEGER - 1),
  YE = E.string().refine(
    (e) => Number.isFinite(Date.parse(e)) && new Date(e).toISOString() === e,
  ),
  q$ = E.strictObject({
    version: E.number().int().min(1).max(xs.versions),
    definition: JE,
    definitionHash: Zi,
    createdAt: YE,
  }),
  W3 = E.array(
    E.strictObject({
      id: zo,
      name: Dt,
      versions: E.array(q$).min(1).max(xs.versions),
    }),
  ).max(xs.recipes),
  KE = E.strictObject({
    name: Dt,
    types: E.array(E.enum(["blank", "number", "string", "boolean"])).max(4),
    declaredType: E.string().max(32).optional(),
  }),
  QE = E.strictObject({
    field: Dt,
    numericCount: Is,
    blankCount: Is,
    sum: E.number().finite(),
  }),
  H$ = E.strictObject({
    index: E.number().int().min(1).max(20),
    id: zo,
    kind: E.string().max(32),
    status: E.enum(["complete", "failed", "skipped"]),
    inputSchema: E.array(KE).max(50),
    inputRows: Is,
    outputSchema: E.array(KE).max(50).optional(),
    outputRows: Is.optional(),
    inputHash: Zi.optional(),
    outputHash: Zi.optional(),
    controls: E.array(QE).max(50).optional(),
    error: E.string().max(4e3).optional(),
    details: E.record(E.string(), E.unknown()).optional(),
  }),
  B3 = E.strictObject({
    engine: E.literal(XE),
    recipeId: zo,
    recipeName: Dt,
    version: E.number().int().min(1).max(20),
    definition: JE,
    definitionHash: Zi,
    sourceWorkbookId: zo,
    sourceRevision: Is,
    sourceHash: Zi,
    outputHash: Zi,
    outputRows: E.number().int().min(0).max(1e3),
    outputRange: E.string().min(1).max(50),
    at: YE,
    steps: E.array(H$).max(20),
    controls: E.array(QE).max(50),
    lineage: E.record(
      E.string().regex(/^[1-9]\d*$/),
      E.array(E.string().min(1).max(200)).max(4e3),
    ),
  });
var ks = (e) => e.replace(/^_xlfn\.(?:_xlws\.)?/i, "").toUpperCase(),
  vn = (e) => e.replace(/^_xlpm\./i, "").toLowerCase();
function K$(e, t = false) {
  return (
    (e = e.replace(/^_xlpm\./i, "")),
    /^[A-Za-z_\\][A-Za-z0-9_.\\]{0,254}$/.test(e) &&
      !/^(?:[A-Za-z]{1,3}[1-9]\d*|R\d+C\d+|R|C)$/i.test(e) &&
      !/^__LG/i.test(e) &&
      (!t || !e.includes("."))
  );
}
function Gi(e) {
  let t = e.startsWith("=") ? 1 : 0,
    r = () => {
      throw Error("Unbalanced or malformed formula syntax.");
    },
    o = (l) => {
      for (t++; t < e.length;)
        if (e[t++] === l)
          if (e[t] === l) t++;
          else return;
      r();
    },
    n = (l, s) => {
      let c = 0;
      do {
        let u = e[t];
        if (u === '"' || (u === "'" && l !== "[")) {
          o(u);
          continue;
        }
        (u === l && c++, u === s && c--, t++);
      } while (t < e.length && c);
      c && r();
    },
    i = (l) => {
      if (l > 128) throw Error("#LIMIT!");
      let s = [];
      for (; t < e.length && ![",", ")"].includes(e[t]);) {
        let c = t,
          u,
          f = e[t];
        if (f === '"') o(f);
        else if (f === "{") n("{", "}");
        else if (f === "[") n("[", "]");
        else if (f === "(") {
          t++;
          let d = i(l + 1);
          (e[t++] !== ")" && r(),
            (u = { kind: "group", body: d, start: c, end: t }));
        } else {
          let d,
            p = /^(?:'((?:[^']|'')+)'|([A-Za-z0-9_\\][A-Za-z0-9_.\\]*))!/.exec(
              e.slice(t),
            );
          p && ((d = p[1]?.replace(/''/g, "'") ?? p[2]), (t += p[0].length));
          let m =
            /^(?:#(?:REF!|NAME\?|VALUE!|DIV\/0!|N\/A|NUM!|NULL!|SPILL!|CALC!|CYCLE!|LIMIT!|BLOCKED!|ERROR!)|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}(?![\w.])|\$?[1-9]\d*:\$?[1-9]\d*(?![\w.])|\$?[A-Za-z]{1,3}\$?[1-9]\d*(?![\w.\\])|(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?)/i.exec(
              e.slice(t),
            );
          if (m) t += m[0].length;
          else {
            let h = /^[A-Za-z_\\][A-Za-z0-9_.\\]*/.exec(e.slice(t));
            h
              ? ((t += h[0].length),
                e[t] === "["
                  ? n("[", "]")
                  : (u = {
                      kind: "symbol",
                      name: h[0],
                      qualifier: d,
                      start: c,
                      end: t,
                    }))
              : ((d || f === "'" || f === "}") && r(), t++);
          }
        }
        if (
          ((u ??= { kind: "text", text: e.slice(c, t), start: c, end: t }),
          u.kind !== "text")
        )
          for (;;) {
            let d = /^\s*\(/.exec(e.slice(t));
            if (!d) break;
            t += d[0].length;
            let p = [];
            if (e[t] !== ")") for (; p.push(i(l + 1)), e[t] === ",";) t++;
            (e[t++] !== ")" && r(),
              (u = { kind: "call", callee: u, args: p, start: c, end: t }));
          }
        s.push(u);
      }
      return s;
    },
    a = i(0);
  return (t !== e.length && r(), a);
}
var qi = (e) => e.filter((t) => t.kind !== "text" || t.text.trim());
function Hi(e, t = false) {
  let r = qi(e),
    o = r[0];
  if (r.length !== 1 || o?.kind !== "symbol" || o.qualifier || !K$(o.name, t))
    throw Error(
      `Invalid ${t ? "LAMBDA parameter" : "LET variable"}: use a name, not a cell address${t ? ", without a period" : ""}.`,
    );
  return o;
}
function gu(e) {
  if (e.callee.kind !== "symbol" || e.callee.qualifier) return;
  let t = ks(e.callee.name);
  if (t !== "LET" && t !== "LAMBDA") return;
  if (
    t === "LET" &&
    (e.args.length < 3 || e.args.length > 253 || e.args.length % 2 !== 1)
  )
    throw Error(
      "LET requires 1\u2013126 name/value pairs and a final calculation.",
    );
  if (t === "LAMBDA" && (e.args.length < 1 || e.args.length > 254))
    throw Error("LAMBDA requires a calculation and at most 253 parameters.");
  if (e.args.some((o) => !qi(o).length))
    throw Error(`${t} arguments cannot be empty.`);
  let r = new Set();
  for (let o = 0; o < e.args.length - 1; o += t === "LET" ? 2 : 1) {
    let n = vn(Hi(e.args[o], t === "LAMBDA").name);
    if (r.has(n)) throw Error(`Duplicate ${t} binding: ${n}.`);
    r.add(n);
  }
  return t;
}
function Pn(e, t, r = new Set()) {
  for (let o of e) {
    if (
      (o.kind === "symbol" && t(o, r, false, false),
      o.kind === "group" && Pn(o.body, t, r),
      o.kind !== "call")
    )
      continue;
    let n = gu(o);
    if (
      (o.callee.kind === "symbol"
        ? t(o.callee, r, true, false)
        : Pn([o.callee], t, r),
      !n)
    ) {
      o.args.forEach((a) => Pn(a, t, r));
      continue;
    }
    let i = new Set(r);
    for (let a = 0; a < o.args.length - 1; a += n === "LET" ? 2 : 1) {
      let l = Hi(o.args[a], n === "LAMBDA");
      (t(l, i, false, true),
        n === "LET" && Pn(o.args[a + 1], t, i),
        i.add(vn(l.name)));
    }
    Pn(o.args.at(-1), t, i);
  }
}
function ex(e) {
  try {
    let t = qi(Gi(e)),
      r = t[0];
    return t.length !== 1 || r.kind !== "call" || gu(r) !== "LAMBDA"
      ? undefined
      : r.args
          .slice(0, -1)
          .map((o) => Hi(o, true).name.replace(/^_xlpm\./i, ""));
  } catch {
    return;
  }
}
var vu = [
  "ABS",
  "ACOS",
  "ACOSH",
  "ACOT",
  "ACOTH",
  "ADDRESS",
  "AND",
  "ARABIC",
  "AREAS",
  "ASC",
  "ASIN",
  "ASINH",
  "ATAN",
  "ATAN2",
  "ATANH",
  "AVEDEV",
  "AVERAGE",
  "AVERAGEA",
  "AVERAGEIF",
  "BAHTTEXT",
  "BASE",
  "BESSELI",
  "BESSELJ",
  "BESSELK",
  "BESSELY",
  "BETA.DIST",
  "BETA.INV",
  "BIN2DEC",
  "BIN2HEX",
  "BIN2OCT",
  "BINOM.DIST",
  "BINOM.DIST.RANGE",
  "BINOM.INV",
  "BITAND",
  "BITLSHIFT",
  "BITOR",
  "BITRSHIFT",
  "BITXOR",
  "CEILING",
  "CEILING.MATH",
  "CEILING.PRECISE",
  "CHAR",
  "CHISQ.DIST",
  "CHISQ.DIST.RT",
  "CHISQ.INV",
  "CHISQ.INV.RT",
  "CHISQ.TEST",
  "CLEAN",
  "CODE",
  "COLUMN",
  "COLUMNS",
  "COMBIN",
  "COMBINA",
  "COMPLEX",
  "CONCAT",
  "CONCATENATE",
  "CONFIDENCE.NORM",
  "CONFIDENCE.T",
  "CORREL",
  "COS",
  "COSH",
  "COT",
  "COTH",
  "COUNT",
  "COUNTIF",
  "COVARIANCE.P",
  "COVARIANCE.S",
  "CSC",
  "CSCH",
  "DATE",
  "DATEDIF",
  "DATEVALUE",
  "DAY",
  "DAYS",
  "DAYS360",
  "DBCS",
  "DEC2BIN",
  "DEC2HEX",
  "DEC2OCT",
  "DECIMAL",
  "DEGREES",
  "DELTA",
  "DEVSQ",
  "DOLLAR",
  "EDATE",
  "ENCODEURL",
  "EOMONTH",
  "ERF",
  "ERFC",
  "ERROR.TYPE",
  "EVEN",
  "EXACT",
  "EXP",
  "EXPON.DIST",
  "F.DIST",
  "F.DIST.RT",
  "F.INV",
  "F.INV.RT",
  "F.TEST",
  "FACT",
  "FACTDOUBLE",
  "FALSE",
  "FIND",
  "FINDB",
  "FISHER",
  "FISHERINV",
  "FIXED",
  "FLOOR",
  "FLOOR.MATH",
  "FLOOR.PRECISE",
  "FORECAST",
  "FORECAST.LINEAR",
  "FREQUENCY",
  "GAMMA",
  "GAMMA.DIST",
  "GAMMA.INV",
  "GAMMALN",
  "GAMMALN.PRECISE",
  "GAUSS",
  "GCD",
  "GEOMEAN",
  "GESTEP",
  "GROWTH",
  "HARMEAN",
  "HEX2BIN",
  "HEX2DEC",
  "HEX2OCT",
  "HLOOKUP",
  "HOUR",
  "HYPGEOM.DIST",
  "IF",
  "IFERROR",
  "IFNA",
  "IFS",
  "IMABS",
  "IMAGINARY",
  "IMARGUMENT",
  "IMCONJUGATE",
  "IMCOS",
  "IMCOSH",
  "IMCOT",
  "IMCSC",
  "IMCSCH",
  "IMDIV",
  "IMEXP",
  "IMLN",
  "IMLOG10",
  "IMLOG2",
  "IMPOWER",
  "IMPRODUCT",
  "IMREAL",
  "IMSEC",
  "IMSECH",
  "IMSIN",
  "IMSINH",
  "IMSQRT",
  "IMSUB",
  "IMSUM",
  "IMTAN",
  "INDEX",
  "INT",
  "INTERCEPT",
  "ISBLANK",
  "ISERR",
  "ISERROR",
  "ISEVEN",
  "ISLOGICAL",
  "ISNA",
  "ISNONTEXT",
  "ISNUMBER",
  "ISO.CEILING",
  "ISOWEEKNUM",
  "ISREF",
  "ISTEXT",
  "KURT",
  "LCM",
  "LEFT",
  "LEFTB",
  "LN",
  "LOG",
  "LOG10",
  "LOGNORM.DIST",
  "LOGNORM.INV",
  "LOWER",
  "MDETERM",
  "MID",
  "MIDB",
  "MINUTE",
  "MMULT",
  "MOD",
  "MONTH",
  "MROUND",
  "MULTINOMIAL",
  "MUNIT",
  "N",
  "NA",
  "NEGBINOM.DIST",
  "NETWORKDAYS",
  "NETWORKDAYS.INTL",
  "NORM.DIST",
  "NORM.INV",
  "NORM.S.DIST",
  "NORM.S.INV",
  "NOT",
  "NOW",
  "NUMBERVALUE",
  "OCT2BIN",
  "OCT2DEC",
  "OCT2HEX",
  "ODD",
  "OR",
  "PHI",
  "PI",
  "POISSON.DIST",
  "POWER",
  "PRODUCT",
  "PROPER",
  "QUOTIENT",
  "RADIANS",
  "RAND",
  "RANDBETWEEN",
  "REPLACE",
  "REPLACEB",
  "REPT",
  "RIGHT",
  "RIGHTB",
  "ROMAN",
  "ROUND",
  "ROUNDDOWN",
  "ROUNDUP",
  "ROW",
  "ROWS",
  "SEARCH",
  "SEARCHB",
  "SEC",
  "SECH",
  "SECOND",
  "SERIESSUM",
  "SIGN",
  "SIN",
  "SINH",
  "SQRT",
  "SQRTPI",
  "STANDARDIZE",
  "SUM",
  "SUMIF",
  "SUMPRODUCT",
  "SUMSQ",
  "SUMX2MY2",
  "SUMX2PY2",
  "SUMXMY2",
  "T",
  "T.DIST",
  "T.DIST.2T",
  "T.DIST.RT",
  "T.INV",
  "T.INV.2T",
  "TAN",
  "TANH",
  "TEXT",
  "TIME",
  "TIMEVALUE",
  "TODAY",
  "TRANSPOSE",
  "TRIM",
  "TRUE",
  "TRUNC",
  "TYPE",
  "UNICHAR",
  "UNICODE",
  "VLOOKUP",
  "WEEKDAY",
  "WEEKNUM",
  "WEIBULL.DIST",
  "WORKDAY",
  "WORKDAY.INTL",
  "XOR",
  "YEAR",
  "YEARFRAC",
];
var X$ = Object.freeze({
    extensions: 8,
    functions: 40,
    calls: 100,
    inputCells: 2e3,
    outputCells: 2e3,
    metadataCharacters: 2e6,
    timeoutMs: 15e3,
    maxTimeoutMs: 6e4,
  }),
  As = (e) =>
    E.string()
      .min(1)
      .max(e)
      .refine((t) => !/[\u0000-\u001f\u007f]/.test(t)),
  yu = E.string()
    .regex(/^[A-Za-z][A-Za-z0-9_-]{0,63}$/)
    .refine((e) => !["constructor", "prototype", "__proto__"].includes(e)),
  J$ = E.enum(["number", "string", "boolean", "scalar"]),
  bu = E.strictObject({ type: J$, nullable: E.boolean().optional() }),
  tx = E.strictObject({
    type: E.literal("matrix"),
    items: bu,
    maxRows: E.number().int().min(1).max(100),
    maxColumns: E.number().int().min(1).max(100),
  }),
  Y$ = E.strictObject({
    type: E.literal("entity"),
    fields: E.record(yu, bu).refine(
      (e) => Object.keys(e).length > 0 && Object.keys(e).length <= 20,
    ),
  }),
  Q$ = E.union([bu, tx, Y$]),
  eO = E.enum([
    "read-arguments",
    "return-results",
    "linked-data",
    "remote-compute",
  ]),
  tO = E.strictObject({
    schemaVersion: E.literal(1),
    id: yu,
    label: As(120),
    version: E.string()
      .regex(/^\d+\.\d+\.\d+(?:-[A-Za-z0-9.-]+)?$/)
      .max(64),
    namespace: E.string()
      .regex(/^[A-Z][A-Z0-9_]{1,23}$/)
      .refine(
        (e) => !["SQL", "REGISTER", "__LG", "_XLFN", "_XLUDF"].includes(e),
      ),
    runtime: E.enum(["host", "office-addin", "python", "r", "linked-data"]),
    permissions: E.array(eO).min(2).max(4),
    recalculation: E.strictObject({
      mode: E.literal("manual"),
      invalidateOn: E.literal("argument-values"),
      volatile: E.literal(false),
    }),
    functions: E.array(
      E.strictObject({
        id: E.string().regex(/^[A-Z][A-Z0-9_]{0,39}$/),
        description: As(300),
        parameters: E.array(
          E.strictObject({ name: yu, schema: E.union([bu, tx]) }),
        ).max(16),
        result: Q$,
      }),
    )
      .min(1)
      .max(20),
  });
function rO(e) {
  let t = tO.parse(e);
  if (
    new Set(t.permissions).size !== t.permissions.length ||
    !t.permissions.includes("read-arguments") ||
    !t.permissions.includes("return-results")
  )
    throw Error(
      "Extensions must declare distinct read-arguments and return-results permissions.",
    );
  if (t.runtime !== "host" && !t.permissions.includes("remote-compute"))
    throw Error("External runtimes must declare remote-compute.");
  if (new Set(t.functions.map((r) => r.id)).size !== t.functions.length)
    throw Error("Duplicate custom function names.");
  for (let r of t.functions) {
    if (vu.includes(`${t.namespace}.${r.id}`))
      throw Error("An extension cannot shadow a built-in function.");
    if (new Set(r.parameters.map((o) => o.name)).size !== r.parameters.length)
      throw Error("Duplicate extension parameter names.");
    if (r.result.type === "entity" && !t.permissions.includes("linked-data"))
      throw Error("Entity results require the linked-data permission.");
  }
  return t;
}
function nO(e, t) {
  let r = (o, n) => {
    if (o === null && n.nullable) return null;
    if (
      (typeof o == "number" &&
        Number.isFinite(o) &&
        (n.type === "number" || n.type === "scalar")) ||
      (typeof o == "string" &&
        o.length <= 32767 &&
        (n.type === "string" || n.type === "scalar")) ||
      (typeof o == "boolean" && (n.type === "boolean" || n.type === "scalar"))
    )
      return o;
    throw Error("Extension value does not match its declared schema.");
  };
  if (t.type === "matrix") {
    if (
      !Array.isArray(e) ||
      !e.length ||
      e.length > t.maxRows ||
      !Array.isArray(e[0]) ||
      !e[0].length ||
      e[0].length > t.maxColumns ||
      e.length * e[0].length > X$.outputCells
    )
      throw Error("Extension matrix exceeds its declared shape or cell limit.");
    let o = e[0].length;
    return e.map((n) => {
      if (!Array.isArray(n) || n.length !== o)
        throw Error("Extension matrix must be rectangular.");
      return n.map((i) => r(i, t.items));
    });
  }
  if (t.type === "entity") {
    let o = E.strictObject({
      kind: E.literal("entity"),
      id: As(200),
      label: As(300),
      fields: E.record(yu, E.unknown()),
      source: As(1e3),
      asOf: E.iso.datetime(),
    }).parse(e);
    if (
      Object.keys(o.fields).sort().join("|") !==
      Object.keys(t.fields).sort().join("|")
    )
      throw Error(
        "Linked entity fields must match the declared schema exactly.",
      );
    return {
      ...o,
      fields: Object.fromEntries(
        Object.entries(t.fields).map(([n, i]) => [n, r(o.fields[n], i)]),
      ),
    };
  }
  return r(e, t);
}
var oO = (e) => Mn(rO(e)),
  iO = (e, t, r, o, n, i) => Mn([e, t, r, o, n, i]);
function R0(e) {
  return (e.extensions?.installed ?? []).flatMap((t) =>
    t.manifest.functions.map((r) => ({
      entry: t,
      fn: r,
      name: `${t.manifest.namespace}.${r.id}`,
    })),
  );
}
var rx = new WeakMap();
function nx(e, t) {
  rx.set(e, structuredClone(t));
}
function wu(e) {
  return rx.get(e) ?? [];
}
function aO(e, t, r, o) {
  return wu(e).some(
    (n) => n.id === t && n.manifestHash === r && n.sheetIds.includes(o),
  );
}
function ox(e, t, r, o, n) {
  let i = R0(e).find((f) => f.name === t);
  if (!i) return { error: "#NAME?" };
  let a = oO(i.entry.manifest);
  if (!i.entry.enabled || !aO(e, i.entry.manifest.id, a, o))
    return { error: "#NAME?" };
  if (r.length !== i.fn.parameters.length) return { error: "#VALUE!" };
  let l;
  try {
    l = r.map((f, d) => {
      let p = i.fn.parameters[d].schema;
      return nO(p.type === "matrix" && !Array.isArray(f) ? [[f]] : f, p);
    });
  } catch {
    return { error: "#VALUE!" };
  }
  let s = String(e.sheets.find((f) => f.id === o)?.cells[n]?.raw ?? ""),
    c = iO(o, n, s, t, l, a),
    u = e.extensions?.caches.find((f) => f.key === c);
  return u &&
    !wu(e).some(
      (f) => f.id === i.entry.manifest.id && f.scopeHash === u.scopeHash,
    )
    ? { error: "#NAME?" }
    : u
      ? { value: u.result }
      : { error: "#PENDING!" };
}
var T0 = Object.freeze({
  defaultMemoryBudgetMiB: 32,
  maxMemoryBudgetMiB: 128,
  defaultWorkBudget: 1e6,
  maxWorkBudget: 1e7,
  maxNesting: 64,
  memoryUnit:
    "Cumulative accounted working data: UTF-16 text, 8-byte scalars and container allowances, including repeated intermediate allocations. This is not a browser heap or process RSS limit.",
  timing:
    "Isolated calculations terminate their worker on deadline or cancellation. Delivery of timer events is subject to host scheduling; synchronous APIs use cooperative checks.",
  scope:
    "Calculation inputs, intermediate values and results. The existing committed workbook, UI, parser code and runtime overhead are outside the working-data budget.",
});
function Wo() {
  return typeof performance < "u" ? performance.now() : Date.now();
}
function C0(e) {
  return {
    memoryBudgetBytes:
      (e?.memoryBudgetMiB ?? T0.defaultMemoryBudgetMiB) * 1024 * 1024,
    workBudget: e?.workBudget ?? T0.defaultWorkBudget,
  };
}
function Eu(e, t) {
  let r = 0,
    o = (n, i) => {
      if (!(r > t)) {
        if (i > 80) {
          r = t + 1;
          return;
        }
        if (typeof n == "string") r += 24 + n.length * 2;
        else if (!n || typeof n != "object") r += 8;
        else if (Array.isArray(n)) {
          r += 32 + n.length * 8;
          for (let a of n) if ((o(a, i + 1), r > t)) break;
        } else {
          r += 64;
          for (let a in n) {
            if (!Object.prototype.hasOwnProperty.call(n, a)) continue;
            r += 32 + a.length * 2;
            let l = Object.getOwnPropertyDescriptor(n, a);
            if (!("value" in l)) {
              r = t + 1;
              break;
            }
            if ((o(l.value, i + 1), r > t)) break;
          }
        }
      }
    };
  return (o(e, 0), r);
}
var ix = new WeakMap();
function S0(e, t) {
  let r = 1024,
    o = (n) => {
      r += Eu(n, Math.max(0, t - r));
    };
  (o(e.definedNames),
    o(e.workbookLinks),
    o(e.extensions),
    o(e.calculation),
    o(e.datePolicy));
  for (let n of e.sheets) {
    o([
      n.id,
      n.name,
      n.tables,
      n.merges,
      n.hiddenRows,
      n.hiddenCols,
      n.outlineGroups,
    ]);
    let i = dn(n.cells),
      a = i ? ix.get(n.cells) : undefined;
    if (a !== undefined) r += a;
    else {
      let l = r;
      for (let s in n.cells)
        if (((r += 128 + 2 * s.length), o(n.cells[s].raw), r > t)) return r;
      i && ix.set(n.cells, r - l);
    }
    if (r > t) return r;
  }
  return r;
}
function xu(e) {
  let t = C0(e.calculation),
    r = S0(e, t.memoryBudgetBytes),
    o = 0,
    n,
    i,
    a = 0,
    l = 0,
    s = 0,
    c = () => (
      !n &&
        r > t.memoryBudgetBytes &&
        ((n = "#LIMIT!"), (i = "Working-data memory budget exceeded")),
      !n &&
        o > t.workBudget &&
        ((n = "#LIMIT!"), (i = "Calculation work budget exceeded")),
      !n &&
        l + (s ? Wo() - a : 0) > (e.calculation?.timeBudgetMs ?? 2500) &&
        ((n = "#TIMEOUT!"),
        (i = "Cooperative calculation time budget exceeded")),
      n
    );
  return {
    begin() {
      s++ === 0 && (a = Wo());
    },
    end() {
      s > 0 && --s === 0 && (l += Wo() - a);
    },
    check: c,
    work(u = 1) {
      return (
        n ||
        (o + Math.max(0, u) > t.workBudget
          ? ((n = "#LIMIT!"), (i = "Calculation work budget exceeded"), n)
          : ((o += Math.max(0, u)), c()))
      );
    },
    reserve(u) {
      return n || ((r += Math.max(0, u)), c());
    },
    value(u) {
      return n || ((r += Eu(u, Math.max(0, t.memoryBudgetBytes - r))), c());
    },
    report: () => ({
      ...t,
      accountedBytes: r,
      workUnits: o,
      ...(n ? { stopped: n, reason: i } : {}),
    }),
  };
}
function Ns(e) {
  let t = e.replace(/"(?:[^"]|"")*"|'(?:[^']|'')*'/g, ""),
    r = 0;
  for (let o of t) {
    if (((o === "(" || o === "{") && r++, r > T0.maxNesting)) return false;
    (o === ")" || o === "}") && r--;
  }
  return true;
}
var Iu = Object.freeze({
    mode: "automatic",
    iterative: false,
    maxIterations: 100,
    maxChange: 0.001,
    timeBudgetMs: 2500,
  }),
  sO = (e) =>
    !!e &&
    typeof e == "object" &&
    !Array.isArray(e) &&
    [Object.prototype, null].includes(Object.getPrototypeOf(e)),
  lO = () => {
    throw Error(
      "Invalid calculation settings. Choose Automatic or Manual, explicit iteration, 1\u20131000 iterations, convergence from 0 to 1, a 1\u201310000 ms time budget, 1\u2013128 MiB working data and 1000\u201310000000 work units.",
    );
  };
function _0(e) {
  return !sO(e) ||
    Object.keys(e).some(
      (t) => ![...Object.keys(Iu), "memoryBudgetMiB", "workBudget"].includes(t),
    ) ||
    typeof e.mode != "string" ||
    !["automatic", "manual"].includes(e.mode) ||
    typeof e.iterative != "boolean" ||
    !Number.isInteger(e.maxIterations) ||
    Number(e.maxIterations) < 1 ||
    Number(e.maxIterations) > 1e3 ||
    typeof e.maxChange != "number" ||
    !Number.isFinite(e.maxChange) ||
    e.maxChange < 0 ||
    e.maxChange > 1 ||
    !Number.isInteger(e.timeBudgetMs) ||
    Number(e.timeBudgetMs) < 1 ||
    Number(e.timeBudgetMs) > 1e4 ||
    (e.memoryBudgetMiB !== undefined &&
      (!Number.isInteger(e.memoryBudgetMiB) ||
        Number(e.memoryBudgetMiB) < 1 ||
        Number(e.memoryBudgetMiB) > 128)) ||
    (e.workBudget !== undefined &&
      (!Number.isInteger(e.workBudget) ||
        Number(e.workBudget) < 1e3 ||
        Number(e.workBudget) > 1e7))
    ? lO()
    : {
        mode: e.mode,
        iterative: e.iterative,
        maxIterations: Number(e.maxIterations),
        maxChange: e.maxChange,
        timeBudgetMs: Number(e.timeBudgetMs),
        ...(e.memoryBudgetMiB !== undefined
          ? { memoryBudgetMiB: Number(e.memoryBudgetMiB) }
          : {}),
        ...(e.workBudget !== undefined
          ? { workBudget: Number(e.workBudget) }
          : {}),
      };
}
var Rs = (e = "1900") => (e === "1904" ? 1462 : 0),
  Ki = (e = "1900") => 2958465 - Rs(e),
  $0 = Object.freeze({
    calendar: "gregorian",
    dateSystem: "1900",
    textDateOrder: "MDY",
    timeZone: "UTC",
  }),
  Xi = Object.freeze({
    maxSerial: 2958465,
    holidays: 1e3,
    businessDays: 366e3,
  });
var ku = 864e5,
  sx = Date.UTC(1899, 11, 31);
function Ji(e = $0) {
  let t = e;
  if (
    !t ||
    typeof t != "object" ||
    Array.isArray(t) ||
    Object.keys(t).some(
      (r) =>
        !["calendar", "dateSystem", "textDateOrder", "timeZone"].includes(r),
    ) ||
    t.calendar !== "gregorian" ||
    !["1900", "1904"].includes(t.dateSystem) ||
    !["MDY", "DMY"].includes(t.textDateOrder) ||
    typeof t.timeZone != "string" ||
    t.timeZone.length > 100 ||
    !/^[A-Za-z_]+(?:\/[A-Za-z0-9_+.-]+)*$/.test(t.timeZone)
  )
    throw Error(
      "Unsupported date policy. Use the Gregorian calendar, 1900 or 1904 serial system, MDY or DMY text dates, and a valid IANA time zone (or UTC).",
    );
  try {
    let r = new Intl.DateTimeFormat("en-US", {
      timeZone: t.timeZone,
    }).resolvedOptions().timeZone;
    return {
      calendar: "gregorian",
      dateSystem: t.dateSystem,
      textDateOrder: t.textDateOrder,
      timeZone: r,
    };
  } catch {
    throw Error(
      `Unsupported time zone: ${t.timeZone}. Use an IANA time zone or UTC.`,
    );
  }
}
var Au = (e) => e === 1900 || (e % 4 === 0 && (e % 100 !== 0 || e % 400 === 0)),
  _r = (e, t) =>
    t === 2 ? (Au(e) ? 29 : 28) : [4, 6, 9, 11].includes(t) ? 30 : 31;
function Yi(e, t, r, o = "1900") {
  if (o === "1904") return (Date.UTC(e, t - 1, r) - Date.UTC(1904, 0, 1)) / ku;
  if (e === 1900 && t === 2 && r === 29) return 60;
  let n = Date.UTC(e, t - 1, r);
  return (n - sx) / ku + (n >= Date.UTC(1900, 2, 1) ? 1 : 0);
}
function Nu(e, t = "1900") {
  let r = Math.floor(e);
  if (t === "1900" && r === 0) return { year: 1900, month: 1, day: 0 };
  if (t === "1900" && r === 60) return { year: 1900, month: 2, day: 29 };
  let o = new Date(
    t === "1904"
      ? Date.UTC(1904, 0, 1) + r * ku
      : sx + (r - (r > 60 ? 1 : 0)) * ku,
  );
  return {
    year: o.getUTCFullYear(),
    month: o.getUTCMonth() + 1,
    day: o.getUTCDate(),
  };
}
function ax(e) {
  let t = /^(\d{1,2}):(\d{2})(?::(\d{2}(?:\.\d{1,9})?))?\s*(AM|PM)?$/i.exec(e);
  if (!t) return;
  let r = Number(t[1]),
    o = Number(t[2]),
    n = Number(t[3] ?? 0),
    i = t[4]?.toUpperCase();
  if (!(o > 59 || n >= 60 || (i ? r < 1 || r > 12 : r > 23)))
    return (
      i && (r = (r % 12) + (i === "PM" ? 12 : 0)),
      (r * 3600 + o * 60 + n) / 86400
    );
}
function Un(e, t = $0) {
  let r = e.trim(),
    o =
      /^(?:(\d{4})-(\d{1,2})-(\d{1,2})|(\d{1,2})\/(\d{1,2})\/(\d{4}))(?:[ T](.+))?$/.exec(
        r,
      );
  if (!o) return ax(r);
  let n = Number(o[1] ?? o[6]),
    i = Number(o[2] ?? o[t.textDateOrder === "MDY" ? 4 : 5]),
    a = Number(o[3] ?? o[t.textDateOrder === "MDY" ? 5 : 4]);
  if (
    n < (t.dateSystem === "1904" ? 1904 : 1900) ||
    n > 9999 ||
    i < 1 ||
    i > 12 ||
    a < 1 ||
    a > _r(n, i)
  )
    return;
  let l = o[7] ? ax(o[7]) : 0;
  return l === undefined ? undefined : Yi(n, i, a, t.dateSystem) + l;
}
function Ts(e, t = new Date()) {
  let r = Ji(e),
    o = t.toISOString(),
    n = new Intl.DateTimeFormat("en-US-u-ca-gregory-nu-latn", {
      timeZone: r.timeZone,
      year: "numeric",
      month: "2-digit",
      day: "2-digit",
      hour: "2-digit",
      minute: "2-digit",
      second: "2-digit",
      hourCycle: "h23",
    }).formatToParts(t),
    i = (h) => Number(n.find((v) => v.type === h)?.value),
    a = i("year"),
    l = i("month"),
    s = i("day"),
    c = i("hour"),
    u = i("minute"),
    f = i("second"),
    d = t.getUTCMilliseconds(),
    p = Yi(a, l, s, r.dateSystem);
  if (p < 0 || p > Ki(r.dateSystem))
    throw Error(
      "Calculation clock is outside the supported 1900\u20139999 date range.",
    );
  let m = `${a}-${String(l).padStart(2, "0")}-${String(s).padStart(2, "0")}T${String(c).padStart(2, "0")}:${String(u).padStart(2, "0")}:${String(f).padStart(2, "0")}.${String(d).padStart(3, "0")}`;
  return Object.freeze({
    ...r,
    calculatedAt: o,
    localDateTime: m,
    utcOffsetMinutes: (Date.UTC(a, l - 1, s, c, u, f, d) - t.getTime()) / 6e4,
    todaySerial: p,
    nowSerial: p + (c * 3600 + u * 60 + f + d / 1e3) / 86400,
  });
}
var lx = Object.freeze({
  supported: ["RAND", "RANDBETWEEN", "NOW", "TODAY"],
  unsupported: ["RANDARRAY", "OFFSET", "INDIRECT", "CELL", "INFO"],
  refresh:
    "Automatic committed revisions and explicit Calculate Now refresh volatile formulas and their dependents. Repeated reads, selection and rendering reuse the same snapshot. Manual edits leave affected results pending until Calculate Now.",
  random:
    "RAND and RANDBETWEEN use the runtime pseudorandom generator without a workbook seed. They are not cryptographic and are not reproducible audit samples. Separate seeded audit sampling never consumes spreadsheet random state.",
  clock:
    "NOW and TODAY share one captured instant and configured time zone per calculation; no background timer refreshes the workbook.",
  inspection:
    "Diagnostic formula evaluation uses a separate snapshot; replay does not redraw random values. Native reopening starts a new calculation; XLSX formula caches are historical only.",
});
function Qi(e) {
  let t = e.replace(/"(?:[^"]|"")*"|'(?:[^']|'')*'(?=\s*!)/g, "");
  return [
    ...new Set(
      [...t.matchAll(/\b(?:_xlfn\.)?(RAND|RANDBETWEEN|NOW|TODAY)\s*\(/gi)].map(
        (r) => r[1].toUpperCase(),
      ),
    ),
  ];
}
var cO = [
    "DATE",
    "DATEVALUE",
    "DAY",
    "MONTH",
    "YEAR",
    "EDATE",
    "EOMONTH",
    "DATEDIF",
    "DAYS",
    "DAYS360",
    "YEARFRAC",
    "WORKDAY",
    "WORKDAY.INTL",
    "NETWORKDAYS",
    "NETWORKDAYS.INTL",
    "WEEKDAY",
    "WEEKNUM",
    "ISOWEEKNUM",
    "TODAY",
    "NOW",
    "TIME",
    "TIMEVALUE",
    "HOUR",
    "MINUTE",
    "SECOND",
  ],
  O0 = Symbol("not-date-function"),
  uO = {
    DATE: [3, 3],
    DATEVALUE: [1, 1],
    DAY: [1, 1],
    MONTH: [1, 1],
    YEAR: [1, 1],
    EDATE: [2, 2],
    EOMONTH: [2, 2],
    DATEDIF: [3, 3],
    DAYS: [2, 2],
    DAYS360: [2, 3],
    YEARFRAC: [2, 3],
    WORKDAY: [2, 3],
    "WORKDAY.INTL": [2, 4],
    NETWORKDAYS: [2, 3],
    "NETWORKDAYS.INTL": [2, 4],
    WEEKDAY: [1, 2],
    WEEKNUM: [1, 2],
    ISOWEEKNUM: [1, 1],
    TODAY: [0, 0],
    NOW: [0, 0],
    TIME: [3, 3],
    TIMEVALUE: [1, 1],
    HOUR: [1, 1],
    MINUTE: [1, 1],
    SECOND: [1, 1],
  };
function Cs(e, t, r) {
  if (!cO.includes(e)) return O0;
  let o = r.clock.dateSystem,
    n = Ki(o),
    i = (b, x, k) => Yi(b, x, k, o),
    a = (b) => Nu(b, o),
    l = (b) => (((b + Rs(o) + 5) % 7) + 7) % 7,
    s = (b = "#VALUE!") => {
      throw r.issue(b);
    },
    [c, u] = uO[e];
  (t.length < c || t.length > u) && s();
  let f = t.map(r.retrieve),
    d = (b) => {
      if (r.isIssue(b)) throw b;
      return Array.isArray(b) ? (b.length === 1 ? d(b[0]) : s()) : b;
    },
    p = (b) => {
      if (((b = d(b)), b == null)) return 0;
      if (typeof b == "boolean") return Number(b);
      typeof b != "number" &&
        (typeof b != "string" ||
          !/^\s*[+-]?(?:\d+\.?\d*|\.\d+)(?:E[+-]?\d+)?\s*$/i.test(b)) &&
        s();
      let x = Number(b);
      return Number.isFinite(x) ? x : s("#NUM!");
    },
    m = (b) => (Number.isFinite(b) && b >= 0 && b < n + 1 ? b : s("#NUM!")),
    h = (b) => {
      b = d(b);
      let x =
        typeof b == "string" &&
        !/^\s*[+-]?(?:\d+\.?\d*|\.\d+)(?:E[+-]?\d+)?\s*$/i.test(b)
          ? Un(b, r.clock)
          : p(b);
      return x === undefined ? s() : m(x);
    },
    v = (b) => Math.floor(h(b)),
    g = (b, x, k = false) => {
      let w = a(b),
        A = w.year * 12 + w.month - 1 + x,
        I = Math.floor(A / 12),
        N = (((A % 12) + 12) % 12) + 1;
      return (
        (I < (o === "1904" ? 1904 : 1900) || I > 9999) && s("#NUM!"),
        m(i(I, N, k ? _r(I, N) : Math.min(w.day, _r(I, N))))
      );
    };
  if (e === "TODAY") return r.clock.todaySerial;
  if (e === "NOW") return r.clock.nowSerial;
  if (e === "DATE") {
    let b = Math.trunc(p(f[0])),
      x = Math.trunc(p(f[1])),
      k = Math.trunc(p(f[2]));
    ((b < 0 || b >= 1e4) && s("#NUM!"), b < 1900 && (b += 1900));
    let w = b * 12 + x - 1;
    b = Math.floor(w / 12);
    let A = (((w % 12) + 12) % 12) + 1;
    return (
      (b < 1899 || b > 1e4 || Math.abs(k) > Xi.maxSerial) && s("#NUM!"),
      m(i(b, A, 1) + k - 1)
    );
  }
  if (e === "DATEVALUE" || e === "TIMEVALUE") {
    let b = d(f[0]);
    typeof b != "string" && s();
    let x = Un(b, r.clock);
    return (
      x === undefined && s(),
      e === "DATEVALUE" &&
        !/^(?:\d{4}-\d{1,2}-\d{1,2}|\d{1,2}\/\d{1,2}\/\d{4})(?:[ T]|$)/.test(
          b.trim(),
        ) &&
        s(),
      e === "DATEVALUE" ? Math.floor(x) : x - Math.floor(x)
    );
  }
  if (e === "TIME") {
    let [b, x, k] = f.map((w) => Math.trunc(p(w)));
    return (
      [b, x, k].some((w) => w < 0 || w > 32767) && s("#NUM!"),
      ((b * 3600 + x * 60 + k) % 86400) / 86400
    );
  }
  if (["HOUR", "MINUTE", "SECOND"].includes(e)) {
    let b = h(f[0]),
      x = Math.max(1e-9, Number.EPSILON * b * 86400),
      k = Math.floor((b - Math.floor(b)) * 86400 + x) % 86400;
    return e === "HOUR"
      ? Math.floor(k / 3600)
      : e === "MINUTE"
        ? Math.floor(k / 60) % 60
        : k % 60;
  }
  if (["DAY", "MONTH", "YEAR"].includes(e)) return a(v(f[0]))[e.toLowerCase()];
  if (e === "EDATE" || e === "EOMONTH")
    return g(v(f[0]), Math.trunc(p(f[1])), e === "EOMONTH");
  if (e === "DAYS") return v(f[0]) - v(f[1]);
  if (e === "WEEKDAY") {
    let b = v(f[0]),
      x = f[1] == null ? 1 : Math.trunc(p(f[1])),
      k = l(b);
    if (x === 3) return k;
    let w =
      x === 1 || x === 17
        ? 6
        : x === 2 || x === 11
          ? 0
          : x >= 12 && x <= 16
            ? x - 11
            : s("#NUM!");
    return ((k - w + 7) % 7) + 1;
  }
  if (e === "WEEKNUM" || e === "ISOWEEKNUM") {
    let b = v(f[0]),
      x = e === "ISOWEEKNUM" ? 21 : f[1] == null ? 1 : Math.trunc(p(f[1]));
    if (x === 21) {
      let A = b + 3 - l(b),
        I = a(A).year;
      return Math.floor((A - i(I, 1, 1)) / 7) + 1;
    }
    let k =
        x === 1 || x === 17
          ? 6
          : x === 2 || x === 11
            ? 0
            : x >= 12 && x <= 16
              ? x - 11
              : s("#NUM!"),
      w = i(a(b).year, 1, 1);
    return Math.floor((b - w + ((l(w) - k + 7) % 7)) / 7) + 1;
  }
  if (e === "DATEDIF") {
    let b = v(f[0]),
      x = v(f[1]),
      k = String(d(f[2])).toUpperCase();
    b > x && s("#NUM!");
    let w = a(b),
      A = a(x),
      I = (A.year - w.year) * 12 + A.month - w.month - (A.day < w.day ? 1 : 0);
    if (k === "D") return x - b;
    if (k === "Y") return Math.floor(I / 12);
    if (k === "M") return I;
    if (k === "YM") return I % 12;
    if (k === "MD") return x - g(b, I);
    if (k === "YD") {
      let N =
        A.year -
        (A.month < w.month || (A.month === w.month && A.day < w.day) ? 1 : 0);
      return x - i(N, w.month, Math.min(w.day, _r(N, w.month)));
    }
    s("#NUM!");
  }
  let y = (b, x, k) => {
    let w = a(b),
      A = a(x),
      I = w.day,
      N = A.day;
    return (
      k
        ? ((I = Math.min(I, 30)), (N = Math.min(N, 30)))
        : (I === _r(w.year, w.month) && (I = 30),
          N === 31 && I >= 30 && (N = 30),
          w.month === 2 &&
            w.day === _r(w.year, 2) &&
            A.month === 2 &&
            N === _r(A.year, 2) &&
            (N = 30)),
      (A.year - w.year) * 360 + (A.month - w.month) * 30 + N - I
    );
  };
  if (e === "DAYS360") {
    let b = d(f[2] ?? false);
    return (
      typeof b != "boolean" && typeof b != "number" && s(),
      y(v(f[0]), v(f[1]), !!b)
    );
  }
  if (e === "YEARFRAC") {
    let b = v(f[0]),
      x = v(f[1]),
      k = f[2] == null ? 0 : Math.trunc(p(f[2]));
    if (((k < 0 || k > 4) && s("#NUM!"), b > x && ([b, x] = [x, b]), b === x))
      return 0;
    if (k === 0 || k === 4) return y(b, x, k === 4) / 360;
    if (k === 2 || k === 3) return (x - b) / (k === 2 ? 360 : 365);
    let w = a(b),
      A = a(x),
      I = 365,
      N = i(w.year + 1, w.month, Math.min(w.day, _r(w.year + 1, w.month)));
    if (w.year === A.year) I = Au(w.year) ? 366 : 365;
    else if (x <= N)
      for (let R = w.year; R <= A.year; R++)
        Au(R) && i(R, 2, 29) >= b && i(R, 2, 29) <= x && (I = 366);
    else I = (i(A.year + 1, 1, 1) - i(w.year, 1, 1)) / (A.year - w.year + 1);
    return (x - b) / I;
  }
  if (e.startsWith("WORKDAY") || e.startsWith("NETWORKDAYS")) {
    let b = e.startsWith("NETWORK"),
      x = e.endsWith(".INTL"),
      k = v(f[0]),
      w = d(x ? (f[2] ?? 1) : 1),
      A;
    if (typeof w == "string") (/^[01]{7}$/.test(w) || s(), (A = w));
    else {
      let Z = p(w),
        B;
      if (Z >= 1 && Z <= 7 && Number.isInteger(Z))
        B = [(Z + 4) % 7, (Z + 5) % 7];
      else if (Z >= 11 && Z <= 17 && Number.isInteger(Z)) B = [(Z - 5) % 7];
      else return s("#NUM!");
      A = Array.from({ length: 7 }, (X, L) => (B.includes(L) ? "1" : "0")).join(
        "",
      );
    }
    let I = (Z) => (Array.isArray(Z) ? Z.flatMap(I) : [Z]),
      N = f[x ? 3 : 2] === undefined ? [] : I(f[x ? 3 : 2]);
    N.length > Xi.holidays && s("#LIMIT!");
    let R = new Set();
    for (let Z of N)
      Z == null || Z === "" || (typeof Z == "boolean" && s(), R.add(v(Z)));
    let C = [...A].filter((Z) => Z === "0").length,
      T = (Z) => A[l(Z)] === "0",
      _ = (Z, B) => {
        if (Z > B) return 0;
        let X = B - Z + 1,
          L = Math.floor(X / 7) * C;
        for (let U = 0; U < X % 7; U++) T(Z + U) && L++;
        for (let U of R) U >= Z && U <= B && T(U) && L--;
        return L;
      };
    if (b) {
      let Z = v(f[1]);
      return (
        Math.abs(Z - k) > Xi.businessDays && s("#LIMIT!"),
        Z < k ? -_(Z, k) : _(k, Z)
      );
    }
    let P = Math.trunc(p(f[1]));
    if ((C || s(), Math.abs(P) > Xi.businessDays && s("#LIMIT!"), P === 0))
      return k;
    let S = P > 0,
      F = Math.abs(P),
      D = S ? n - k : k,
      Q = (Z) => (S ? _(k + 1, k + Z) : _(k - Z, k - 1));
    Q(D) < F && s("#NUM!");
    let te = 1,
      K = D;
    for (; te < K;) {
      let Z = Math.floor((te + K) / 2);
      Q(Z) >= F ? (K = Z) : (te = Z + 1);
    }
    return k + (S ? te : -te);
  }
  return s();
}
var $r = Object.freeze({
  cashFlows: 1e4,
  periods: 1e4,
  couponPeriods: 1e3,
  evaluations: 2048,
  minRate: -0.999999999,
  maxRate: 1e6,
  relativeResidual: 1e-11,
});
function cx(e, t, r = $r.minRate) {
  let o = Math.log1p(r),
    n = Math.log1p($r.maxRate),
    i = Math.log1p(t),
    a = 0,
    l = [],
    s = (v) => {
      if (v < o || v > n || !Number.isFinite(v) || a >= $r.evaluations) return;
      a++;
      let g = e(Math.expm1(v));
      if (!(
        !Number.isFinite(g.value) ||
        !Number.isFinite(g.scale) ||
        g.scale <= 0
      ))
        return { ...g, normalized: g.value / g.scale };
    },
    c = (v, g) => {
      if (!g || Math.abs(g.normalized) > $r.relativeResidual) return false;
      let y = Math.expm1(v);
      return (
        l.some(
          (b) => Math.abs(b.rate - y) <= 1e-7 * Math.max(1, Math.abs(y)),
        ) ||
          l.push({
            rate: y,
            residual: g.value,
            relativeResidual: Math.abs(g.normalized),
          }),
        true
      );
    };
  if ([0, Math.log1p(0.1), Math.log1p(0.5)].map(s).every((v) => v?.value === 0))
    return {
      status: "indeterminate",
      candidates: l,
      evaluations: a,
      exhaustive: false,
    };
  let f = i;
  for (let v = 0; v < 60; v++) {
    let g = s(f);
    if (!g || c(f, g)) break;
    let y = 1e-5,
      b = s(Math.max(o, f - y)),
      x = s(Math.min(n, f + y));
    if (!b || !x) break;
    let k =
      (x.normalized - b.normalized) / (Math.min(n, f + y) - Math.max(o, f - y));
    if (!Number.isFinite(k) || Math.abs(k) < 1e-15) break;
    let w = f - Math.max(-1, Math.min(1, g.normalized / k));
    if (w < o || w > n || w === f) break;
    f = w;
  }
  let d = [
      ...new Set([
        ...Array.from({ length: 193 }, (v, g) => o + ((n - o) * g) / 192),
        ...[
          -0.9,
          -0.5,
          -0.1,
          0,
          0.01,
          0.025,
          0.05,
          0.075,
          0.1,
          0.15,
          0.2,
          0.3,
          0.5,
          1,
          2,
          10,
          t,
        ]
          .filter((v) => v >= r)
          .map(Math.log1p),
      ]),
    ].sort((v, g) => v - g),
    p;
  for (let v of d) {
    let g = s(v);
    if (!g) {
      p = undefined;
      continue;
    }
    if ((c(v, g), p && p.y.normalized * g.normalized < 0)) {
      let y = p.x,
        b = v,
        x = p.y.normalized;
      for (let k = 0; k < 100; k++) {
        let w = (y + b) / 2,
          A = s(w);
        if (!A || c(w, A) || w === y || w === b) break;
        x * A.normalized < 0 ? (b = w) : ((y = w), (x = A.normalized));
      }
    }
    p = { x: v, y: g };
  }
  l.sort((v, g) => v.rate - g.rate);
  let m = [...l].sort(
      (v, g) =>
        Math.abs(Math.log1p(v.rate) - i) - Math.abs(Math.log1p(g.rate) - i),
    )[0],
    h = a >= $r.evaluations;
  return {
    ...(m && !h ? { selected: m } : {}),
    candidates: l,
    evaluations: a,
    status: h ? "limit" : m ? "converged" : "no-convergence",
    exhaustive: false,
  };
}
var P0 = [
    "PV",
    "FV",
    "PMT",
    "IPMT",
    "PPMT",
    "RATE",
    "NPER",
    "NPV",
    "XNPV",
    "IRR",
    "XIRR",
    "MIRR",
    "SLN",
    "DB",
    "DDB",
    "SYD",
    "PRICE",
    "YIELD",
    "ACCRINT",
    "ACCRINTM",
  ],
  U0 = Symbol("not-financial"),
  fO = {
    PV: [3, 5],
    FV: [3, 5],
    PMT: [3, 5],
    IPMT: [4, 6],
    PPMT: [4, 6],
    RATE: [3, 6],
    NPER: [3, 5],
    NPV: [2, 255],
    XNPV: [3, 3],
    IRR: [1, 2],
    XIRR: [2, 3],
    MIRR: [3, 3],
    SLN: [3, 3],
    DB: [4, 5],
    DDB: [4, 5],
    SYD: [4, 4],
    PRICE: [6, 7],
    YIELD: [6, 7],
    ACCRINT: [6, 8],
    ACCRINTM: [4, 5],
  },
  M0 = (e) => (Array.isArray(e) ? e.flatMap(M0) : [e]),
  Bo = (e) => {
    let t = 0,
      r = 0;
    for (let o of e) {
      let n = o - r,
        i = t + n;
      ((r = i - t - n), (t = i));
    }
    return t;
  },
  dO = (e) => {
    let t = e.filter((r) => r !== 0).map(Math.sign);
    return t.slice(1).filter((r, o) => r !== t[o]).length;
  };
function ux(e, t, r) {
  if (!P0.includes(e)) return U0;
  let o =
      "Cash paid out is negative; cash received is positive. Rates and periods must use the same interval. No automatic cent rounding.",
    n = false,
    i = r.clock.dateSystem,
    a = (w, A, I) => Yi(w, A, I, i),
    l = (w) => Nu(w, i),
    s = (w) => {
      ((n = true), r.report?.({ functionName: e, convention: o, ...w }));
    },
    c = (w, A) => {
      throw (s({ status: "error", error: w, message: A }), r.issue(w));
    },
    u = (w) =>
      Number.isFinite(w)
        ? w
        : c(
            "#NUM!",
            "The financial equation overflowed or has no finite result. Review rates, signs and periods.",
          ),
    f = (w) => {
      if (r.isIssue(w)) throw w;
      return Array.isArray(w)
        ? w.length === 1
          ? f(w[0])
          : c("#VALUE!", "A scalar argument must contain exactly one value.")
        : w;
    },
    d = (w, A) => (
      (w = f(w)),
      w == null
        ? (A ?? 0)
        : typeof w == "boolean"
          ? Number(w)
          : typeof w != "number" &&
              (typeof w != "string" ||
                !/^\s*[+-]?(?:\d+\.?\d*|\.\d+)(?:E[+-]?\d+)?\s*$/i.test(w))
            ? c(
                "#VALUE!",
                "Expected a finite number; nonnumeric text cannot be used as a financial input.",
              )
            : u(Number(w))
    ),
    p = (w) => {
      if (((w = f(w)), w == null || typeof w == "boolean"))
        return c(
          "#VALUE!",
          "A cash-flow or security date is missing or is not a date.",
        );
      let A = typeof w == "string" ? Un(w, r.clock) : d(w);
      return A === undefined || A < (i === "1904" ? 0 : 1) || A >= Ki(i) + 1
        ? c(
            "#VALUE!",
            "Invalid cash-flow or security date. Use a valid Gregorian/workbook-system serial or explicit full-year date text.",
          )
        : Math.floor(A);
    },
    m = (w) => (
      w <= 0 && c("#NUM!", "Periods must be greater than zero."),
      w > $r.periods &&
        c("#LIMIT!", "Financial periods exceed the 10000-period budget."),
      w
    ),
    h = (w) =>
      w > -1 ? w : c("#NUM!", "The rate must be greater than -100%."),
    v = (w) =>
      [0, 1].includes(w)
        ? w
        : c("#NUM!", "Payment timing must be 0 (end) or 1 (beginning)."),
    g = (w, A) => Math.exp(A * Math.log1p(w)),
    y = (w, A) => (w === 0 ? A : Math.expm1(A * Math.log1p(w)) / w),
    b = (w, A, I, N, R) => -(I * g(w, A) + N) / ((1 + w * R) * y(w, A)),
    x = (w, A = false) => {
      let I = M0(w);
      I.length > $r.cashFlows &&
        c("#LIMIT!", "Cash-flow input exceeds 10000 positions.");
      let N = [];
      for (let R of I) {
        if (r.isIssue(R)) throw R;
        typeof R == "number"
          ? N.push(u(R))
          : A &&
            c(
              "#VALUE!",
              "Dated cash flows require one numeric amount for every date; blanks, booleans and text cannot be silently removed.",
            );
      }
      return (
        N.length || c("#NUM!", "No numeric cash flows were supplied."),
        N
      );
    },
    k = (w, A, I, N) => {
      (A < (N ?? $r.minRate) || A > $r.maxRate) &&
        c("#NUM!", "The rate guess is outside the supported search interval.");
      let R = cx(w, A, N),
        C = R.status === "limit" ? "#LIMIT!" : "#NUM!";
      if (
        (s({
          status: R.selected ? "converged" : "error",
          ...(R.selected ? {} : { error: C }),
          guess: A,
          signChanges: I,
          solver: R,
          message: R.selected
            ? `Residual-checked rate. ${R.candidates.length} candidate root(s) found; selected nearest the guess in log-rate space. ${I !== undefined && I > 1 ? "Cash flows change sign more than once; multiple returns may exist. " : ""}The bounded search does not prove uniqueness.`
            : R.status === "indeterminate"
              ? "The equation does not identify a unique rate; no rate has been invented."
              : "No residual-checked rate converged within the supported search interval and evaluation budget. Review signs, timing and guess.",
        }),
        !R.selected)
      )
        throw r.issue(C);
      return R.selected.rate;
    };
  try {
    let [w, A] = fO[e];
    (t.length < w || t.length > A) &&
      c("#VALUE!", `${e} requires ${w}\u2013${A} arguments.`);
    let I = t.map(r.retrieve),
      N;
    if (["PV", "FV", "PMT", "IPMT", "PPMT", "RATE", "NPER"].includes(e))
      if (e === "RATE") {
        let R = m(d(I[0])),
          C = d(I[1]),
          T = d(I[2]),
          _ = d(I[3]),
          P = v(d(I[4]));
        ([T, C, _].every((S) => S === 0) &&
          c("#NUM!", "All amounts are zero; no rate can be determined."),
          (N = k(
            (S) => {
              let F = [T * g(S, R), C * (1 + S * P) * y(S, R), _];
              return { value: Bo(F), scale: Bo(F.map(Math.abs)) };
            },
            d(I[5], 0.1),
          )));
      } else if (e === "NPER") {
        let R = h(d(I[0])),
          C = d(I[1]),
          T = d(I[2]),
          _ = d(I[3]),
          P = v(d(I[4]));
        if (R === 0)
          (C === 0 &&
            c(
              "#DIV/0!",
              "Zero rate and zero payment cannot determine a period count.",
            ),
            (N = -(T + _) / C));
        else {
          let S = C * (1 + R * P),
            F = (S - _ * R) / (T * R + S);
          (F > 0 ||
            c(
              "#NUM!",
              "Cash-flow signs or payment size do not produce a finite real period count.",
            ),
            (N = Math.log(F) / Math.log1p(R)));
        }
        m(N);
      } else if (e === "IPMT" || e === "PPMT") {
        let R = h(d(I[0])),
          C = d(I[1]),
          T = m(d(I[2])),
          _ = d(I[3]),
          P = d(I[4]),
          S = v(d(I[5]));
        (C < 1 || C > T || !Number.isInteger(C)) &&
          c(
            "#NUM!",
            "Interest/principal period must be an integer from 1 through nper.",
          );
        let F = b(R, T, _, P, S),
          D =
            S && C === 1
              ? 0
              : (-(_ * g(R, C - 1) + F * (1 + R * S) * y(R, C - 1)) * R) /
                (1 + R * S);
        N = e === "IPMT" ? D : F - D;
      } else {
        let R = h(d(I[0])),
          C = m(d(I[1])),
          T = d(I[2]),
          _ = d(I[3]),
          P = v(d(I[4]));
        N =
          e === "PMT"
            ? b(R, C, T, _, P)
            : e === "FV"
              ? -(_ * g(R, C) + T * (1 + R * P) * y(R, C))
              : -(_ + T * (1 + R * P) * y(R, C)) / g(R, C);
      }
    else if (["NPV", "XNPV", "IRR", "XIRR", "MIRR"].includes(e)) {
      let R = e.startsWith("X"),
        C = e === "IRR" || e === "XIRR",
        T = x(e === "NPV" ? I.slice(1) : I[C || e === "MIRR" ? 0 : 1], R),
        _ = T.map((S, F) => F + (e === "NPV" ? 1 : 0));
      if (R) {
        let S = M0(I[C ? 1 : 2]);
        S.length !== T.length &&
          c("#NUM!", "Every cash-flow amount must have exactly one date.");
        let F = S.map(p);
        (F.some((D) => D < F[0]) &&
          c("#NUM!", "A cash-flow date precedes the first cash-flow date."),
          (_ = F.map((D) => (D - F[0]) / 365)),
          (o =
            "Dated cash flows use truncated Gregorian/workbook-system dates and actual elapsed days / 365, including leap years. The first date is time zero; later dates may be unsorted or repeated. One numeric amount per date."));
      } else
        o =
          "Periodic cash-flow positions are equally spaced. NPV starts at period 1; IRR/MIRR start at period 0. Numeric zeros retain a period; blanks, text and booleans in cash-flow lists are ignored. No automatic cent rounding.";
      let P = (S) => {
        let F = T.map((D, Q) => D / Math.exp(_[Q] * Math.log1p(S)));
        return { value: Bo(F), scale: Bo(F.map(Math.abs)) };
      };
      if (C || e === "MIRR")
        if (
          ((!T.some((S) => S < 0) || !T.some((S) => S > 0)) &&
            c(
              "#NUM!",
              "At least one positive and one negative cash flow are required.",
            ),
          e === "MIRR")
        ) {
          let S = h(d(I[1])),
            F = h(d(I[2])),
            D = T.length - 1,
            Q = Bo(
              T.map((K, Z) => (K < 0 ? K / Math.exp(Z * Math.log1p(S)) : 0)),
            ),
            te = Bo(
              T.map((K, Z) =>
                K > 0 ? K * Math.exp((D - Z) * Math.log1p(F)) : 0,
              ),
            );
          N = Math.expm1(Math.log(-te / Q) / D);
        } else N = k(P, d(I[R ? 2 : 1], 0.1), dO(T));
      else N = P(h(d(I[0]))).value;
    } else if (["SLN", "SYD", "DB", "DDB"].includes(e)) {
      let R = d(I[0]),
        C = d(I[1]),
        T = m(d(I[2]));
      if (
        ((o =
          "Depreciation uses matching life/period units and no automatic cent rounding. DB rounds its fixed rate to three decimals and prorates first/last years; DDB does not switch to straight line."),
        (R < 0 || C < 0) && c("#NUM!", "Cost and salvage must be nonnegative."),
        e === "SLN")
      )
        N = (R - C) / T;
      else {
        let _ = d(I[3]);
        if (
          ((_ < 1 || _ > T + (e === "DB" ? 1 : 0) || !Number.isInteger(_)) &&
            c(
              "#NUM!",
              "Depreciation period must be a supported whole period within asset life (plus DB partial final year).",
            ),
          e === "SYD")
        )
          N = ((R - C) * (T - _ + 1) * 2) / (T * (T + 1));
        else if (e === "DB") {
          let P = d(I[4], 12);
          if (
            ((!Number.isInteger(T) ||
              P < 1 ||
              P > 12 ||
              !Number.isInteger(P) ||
              _ > T + (P < 12 ? 1 : 0)) &&
              c(
                "#NUM!",
                "DB requires whole life/period and 1\u201312 first-year months; a partial first year permits one partial final year.",
              ),
            R === 0 || C >= R)
          )
            N = 0;
          else {
            let S = Math.round((1 - (C / R) ** (1 / T)) * 1e3) / 1e3,
              F = R;
            N = 0;
            for (let D = 1; D <= _; D++)
              ((N = F * S * (D === 1 ? P / 12 : D > T ? (12 - P) / 12 : 1)),
                (F -= N));
          }
        } else {
          let P = d(I[4], 2);
          P <= 0 && c("#NUM!", "DDB factor must be greater than zero.");
          let S = R;
          N = 0;
          for (let F = 1; F <= _; F++)
            ((N = Math.max(0, Math.min((S * P) / T, S - C))), (S -= N));
        }
      }
    } else {
      let R = e === "ACCRINT" ? 6 : e === "ACCRINTM" ? 4 : 6,
        C = Math.trunc(d(I[R]));
      ([0, 1, 2, 3, 4].includes(C) ||
        c(
          "#NUM!",
          "Day-count basis must be 0 (US 30/360), 1 (actual/actual), 2 (actual/360), 3 (actual/365), or 4 (European 30/360).",
        ),
        (o = `Security basis ${C}; annual coupon/yield rates, regular maturity-anchored coupons at frequency 1/2/4, end-of-month anchoring. PRICE/YIELD are per 100 face value. Gregorian/workbook-system dates; no business-day adjustment or automatic rounding.`));
      let T = (F, D) =>
          C === 0 || C === 4
            ? Cs("DAYS360", [F, D, C === 4], { ...r, retrieve: (Q) => Q })
            : D - F,
        _ = (F, D, Q) => (C === 1 ? D - F : (C === 3 ? 365 : 360) / Q),
        P = (F) => {
          let D = Math.trunc(d(F));
          return [1, 2, 4].includes(D)
            ? D
            : c("#NUM!", "Coupon frequency must be 1, 2 or 4.");
        },
        S = (F, D) => {
          let Q = l(F),
            te = Q.year * 12 + Q.month - 1 + D,
            K = Math.floor(te / 12),
            Z = (((te % 12) + 12) % 12) + 1;
          return a(
            K,
            Z,
            Q.day === _r(Q.year, Q.month)
              ? _r(K, Z)
              : Math.min(Q.day, _r(K, Z)),
          );
        };
      if (e === "PRICE" || e === "YIELD") {
        let F = p(I[0]),
          D = p(I[1]),
          Q = d(I[2]),
          te = d(I[3]),
          K = d(I[4]),
          Z = P(I[5]);
        (F >= D || Q < 0 || K <= 0 || te < 0 || (e === "YIELD" && te === 0)) &&
          c(
            "#NUM!",
            "Require settlement before maturity, nonnegative coupon/yield, and positive redemption/price.",
          );
        let B = 1,
          X = S(D, -12 / Z);
        for (; X > F;)
          (B >= $r.couponPeriods &&
            c("#LIMIT!", "Security exceeds 1000 coupon periods."),
            B++,
            (X = S(D, (-B * 12) / Z)));
        let L = S(D, (-(B - 1) * 12) / Z),
          U = _(X, L, Z),
          Y = T(X, F),
          ie = C === 0 || C === 4 ? U - Y : T(F, L),
          ve = ie / U,
          $ = (100 * Q) / Z,
          W = (se) =>
            B === 1
              ? (K + $) / (1 + (se / Z) * ve) - ($ * Y) / U
              : Bo(
                  Array.from(
                    { length: B },
                    (me, we) => $ / (1 + se / Z) ** (we + ve),
                  ),
                ) +
                K / (1 + se / Z) ** (B - 1 + ve) -
                ($ * Y) / U;
        N =
          e === "PRICE"
            ? W(te)
            : k(
                (se) => {
                  let me = W(se);
                  return { value: me - te, scale: Math.abs(me) + te };
                },
                0.1,
                undefined,
                0,
              );
      } else if (e === "ACCRINTM") {
        let F = p(I[0]),
          D = p(I[1]),
          Q = d(I[2]),
          te = d(I[3], 1e3);
        (F >= D || Q <= 0 || te <= 0) &&
          c(
            "#NUM!",
            "Require issue before settlement, positive coupon rate and positive par.",
          );
        let K = Cs("YEARFRAC", [F, D, C], { ...r, retrieve: (Z) => Z });
        ((N = te * Q * K),
          (o = `Accrued interest at maturity: par \xD7 annual rate \xD7 YEARFRAC basis ${C}, using the published local day-count convention. Gregorian/workbook-system; no automatic rounding.`));
      } else {
        let F = p(I[0]),
          D = p(I[1]),
          Q = p(I[2]),
          te = d(I[3]),
          K = d(I[4], 1e3),
          Z = P(I[5]),
          B = v(d(I[7], 1));
        (F >= Q || D <= F || te <= 0 || K <= 0) &&
          c(
            "#NUM!",
            "Require issue before first interest and settlement, positive rate and positive par.",
          );
        let X = !B && Q > D ? D : F,
          L = Math.floor(
            (l(X).year * 12 + l(X).month - l(D).year * 12 - l(D).month) /
              (12 / Z),
          );
        for (; S(D, (L * 12) / Z) > X;) L--;
        let U = S(D, (L * 12) / Z),
          Y = S(D, ((L + 1) * 12) / Z),
          ie = 0,
          ve = 0;
        for (; U < Q;) {
          ++ve > $r.couponPeriods &&
            c("#LIMIT!", "Accrual exceeds 1000 quasi-coupon periods.");
          let $ = Math.max(X, U),
            W = Math.min(Q, Y);
          (W > $ && (ie += T($, W) / _(U, Y, Z)),
            (U = Y),
            (Y = S(D, ((++L + 1) * 12) / Z)));
        }
        ((N = ((K * te) / Z) * ie),
          (o = `Accrued coupon interest: quasi-coupon periods anchored to first_interest, basis ${C}, frequency ${Z}; calc_method ${B} starts at ${!B && Q > D ? "first_interest" : "issue"}. No business-day adjustment or automatic rounding.`));
      }
    }
    return (
      (N = u(N)),
      n ||
        s({
          status: "calculated",
          message:
            "Calculated under the displayed sign, timing and day-count convention.",
        }),
      N
    );
  } catch (w) {
    throw (
      n ||
        s({
          status: "error",
          error: r.isIssue(w) ? (w.error ?? "#VALUE!") : "#VALUE!",
          message:
            "An input or dependent calculation failed; inspect the source values and dates.",
        }),
      w
    );
  }
}
var yt = {
  maxRows: 1048576,
  maxCols: 16384,
  maxCells: 1e5,
  maxFileBytes: 10485760,
  maxNativeBytes: 52428800,
  maxLLMCells: 500,
  maxPatches: 500,
};
var en = {
  ...yt,
  maxCellTextLength: 32767,
  maxFormulaLength: 4096,
  maxFormulaRangeCells: 2e4,
  maxFormulaDepth: 64,
  maxSpillCells: 1e4,
  maxXlsxExpandedBytes: 50 * 1024 * 1024,
  maxXlsxEntries: 2e3,
  maxXlsxExpansionRatio: 200,
  xlsxRatioThresholdBytes: 1024 * 1024,
  maxXlsxCellRecords: yt.maxCells,
};
function Ru(e, t, r) {
  if (!e.some(Array.isArray)) return t(e);
  let o = e.map((a) =>
      Array.isArray(a) ? (a.every(Array.isArray) ? a : [a]) : [[a]],
    ),
    n = Math.max(...o.map((a) => a.length)),
    i = Math.max(...o.map((a) => a[0]?.length ?? 0));
  return !n ||
    !i ||
    o.some(
      (a) =>
        !a.length ||
        !a[0].length ||
        a.some((l) => l.length !== a[0].length) ||
        (a.length !== n && a.length !== 1) ||
        (a[0].length !== i && a[0].length !== 1),
    )
    ? r("#VALUE!")
    : n * i > en.maxSpillCells
      ? r("#LIMIT!")
      : Array.from({ length: n }, (a, l) =>
          Array.from({ length: i }, (s, c) =>
            t(
              o.map(
                (u) => u[u.length === 1 ? 0 : l][u[0].length === 1 ? 0 : c],
              ),
            ),
          ),
        );
}
var mO = [
    "LEFT",
    "RIGHT",
    "MID",
    "LEN",
    "TRIM",
    "CLEAN",
    "UPPER",
    "LOWER",
    "PROPER",
    "FIND",
    "SEARCH",
    "SUBSTITUTE",
    "REPLACE",
    "VALUE",
    "NUMBERVALUE",
    "TEXT",
    "CONCAT",
    "CONCATENATE",
    "TEXTJOIN",
    "EXACT",
    "TEXTBEFORE",
    "TEXTAFTER",
  ],
  Tu = Symbol("not-text-function"),
  L0 = (e) => (Array.isArray(e) ? e.flatMap(L0) : [e]),
  tn = (e) =>
    e == null ? "" : typeof e == "boolean" ? (e ? "TRUE" : "FALSE") : String(e);
function pO(e) {
  let t =
      /^(?:(\d{4})-(\d{1,2})-(\d{1,2})|(\d{1,2})\/(\d{1,2})\/(\d{4}))(?:[ T](.+))?$/.exec(
        e,
      ),
    r = 0,
    o = e;
  if (t) {
    let c = Number(t[1] ?? t[6]),
      u = Number(t[2] ?? t[4]),
      f = Number(t[3] ?? t[5]);
    if (c < 1900 || c > 9999 || u < 1 || u > 12 || f < 1) return;
    let d = c === 1900 && u === 2 && f === 29,
      p = Date.UTC(c, u - 1, f),
      m = new Date(p);
    if (!d && (m.getUTCMonth() !== u - 1 || m.getUTCDate() !== f)) return;
    if (
      ((r = d
        ? 60
        : (p - Date.UTC(1899, 11, 31)) / 864e5 +
          (p >= Date.UTC(1900, 2, 1) ? 1 : 0)),
      !t[7])
    )
      return r;
    o = t[7];
  }
  let n = /^(\d{1,2}):(\d{2})(?::(\d{2}(?:\.\d+)?))?\s*(AM|PM)?$/i.exec(o);
  if (!n) return;
  let i = Number(n[1]),
    a = Number(n[2]),
    l = Number(n[3] ?? 0),
    s = n[4]?.toUpperCase();
  if (!(a > 59 || l >= 60 || (s ? i < 1 || i > 12 : i > 23)))
    return (
      s && (i = (i % 12) + (s === "PM" ? 12 : 0)),
      r + (i * 3600 + a * 60 + l) / 86400
    );
}
function D0(e, t, r) {
  if (!mO.includes(e)) return Tu;
  let o = (h = "#VALUE!") => {
      throw r.issue(h);
    },
    n = {
      LEFT: [1, 2],
      RIGHT: [1, 2],
      MID: [3, 3],
      LEN: [1, 1],
      TRIM: [1, 1],
      CLEAN: [1, 1],
      UPPER: [1, 1],
      LOWER: [1, 1],
      PROPER: [1, 1],
      FIND: [2, 3],
      SEARCH: [2, 3],
      SUBSTITUTE: [3, 4],
      REPLACE: [4, 4],
      VALUE: [1, 1],
      NUMBERVALUE: [1, 3],
      TEXT: [2, 2],
      CONCAT: [1, 253],
      CONCATENATE: [1, 255],
      TEXTJOIN: [3, 254],
      EXACT: [2, 2],
      TEXTBEFORE: [2, 6],
      TEXTAFTER: [2, 6],
    },
    [i, a] = n[e];
  (t.length < i || t.length > a) && o();
  let l = t.map(r.retrieve),
    s = l.flatMap(L0).find(r.isIssue);
  if (s) return s;
  ((!(e === "CONCAT" || e === "TEXTJOIN") && l.some(Array.isArray)) ||
    (e === "TEXTJOIN" && l.slice(0, 2).some(Array.isArray))) &&
    o();
  let u = (h) => (h.length <= 32767 ? h : o("#LIMIT!")),
    f = (h, v) => {
      if (h === undefined && v !== undefined) return v;
      if (Array.isArray(h) || (typeof h == "string" && !h.trim())) return o();
      let g = Number(h ?? 0);
      return Number.isFinite(g) ? g : o();
    },
    d = (h, v) => Math.trunc(f(h, v)),
    p = (h) =>
      typeof h == "string"
        ? (/^(true|false)$/i.test(h) || o(), h.toLowerCase() === "true")
        : !!h,
    m = tn(l[0]);
  if (e === "LEN") return [...m].length;
  if (e === "TRIM") return u(m.replace(/^ +| +$/g, "").replace(/ +/g, " "));
  if (e === "CLEAN") return u(m.replace(/[\u0000-\u001f]/g, ""));
  if (e === "UPPER") return u(m.toUpperCase());
  if (e === "LOWER") return u(m.toLowerCase());
  if (e === "PROPER")
    return u(
      m
        .toLowerCase()
        .replace(
          /(^|[^\p{L}\p{M}])(\p{L})/gu,
          (h, v, g) => v + g.toUpperCase(),
        ),
    );
  if (e === "EXACT") return m === tn(l[1]);
  if (e === "LEFT" || e === "RIGHT") {
    let h = d(l[1], 1),
      v = [...m];
    return (
      h < 0 && o(),
      (e === "LEFT" ? v.slice(0, h) : h ? v.slice(-h) : []).join("")
    );
  }
  if (e === "MID" || e === "REPLACE") {
    let h = d(l[1]),
      v = d(l[2]),
      g = [...m];
    return (
      (h < 1 || v < 0) && o(),
      e === "MID"
        ? g.slice(h - 1, h - 1 + v).join("")
        : u(g.slice(0, h - 1).join("") + tn(l[3]) + g.slice(h - 1 + v).join(""))
    );
  }
  if (e === "FIND" || e === "SEARCH") {
    let h = [...tn(l[1])],
      v = d(l[2], 1);
    (v < 1 || v > Math.max(1, h.length)) && o();
    let g = [...m],
      y = [];
    for (let k = 0; k < g.length; k++) {
      let w = g[k];
      e === "SEARCH" && w === "~" && ["*", "?", "~"].includes(g[k + 1])
        ? y.push({ char: g[++k], kind: "literal" })
        : y.push({
            char: w,
            kind:
              e === "SEARCH" && w === "*"
                ? "many"
                : e === "SEARCH" && w === "?"
                  ? "one"
                  : "literal",
          });
    }
    y.length * h.length > 2e6 && o("#LIMIT!");
    let b = Array(h.length + 1).fill(true);
    for (let k = y.length - 1; k >= 0; k--) {
      let w = y[k],
        A = Array(h.length + 1).fill(false);
      A[h.length] = w.kind === "many" && b[h.length];
      for (let I = h.length - 1; I >= v - 1; I--)
        A[I] =
          w.kind === "many"
            ? b[I] || A[I + 1]
            : b[I + 1] &&
              (w.kind === "one" ||
                (e === "SEARCH"
                  ? h[I].toLowerCase() === w.char.toLowerCase()
                  : h[I] === w.char));
      b = A;
    }
    let x = b.findIndex((k, w) => w >= v - 1 && k);
    return x < 0 ? o() : x + 1;
  }
  if (e === "SUBSTITUTE") {
    let h = tn(l[1]),
      v = tn(l[2]),
      g = l[3] === undefined ? undefined : d(l[3]);
    if ((g !== undefined && g < 1 && o(), !h)) return m;
    let y = m.split(h),
      b = y.length - 1;
    return (
      m.length +
        (g === undefined ? b : g <= b ? 1 : 0) * (v.length - h.length) >
        32767 && o("#LIMIT!"),
      y.reduce(
        (x, k, w) =>
          w === 0 ? k : x + (g === undefined || g === w ? v : h) + k,
        "",
      )
    );
  }
  if (e === "CONCAT" || e === "CONCATENATE" || e === "TEXTJOIN") {
    let h = e === "TEXTJOIN" ? tn(l[0]) : "",
      v = e === "TEXTJOIN" && p(l[1]),
      g = l
        .slice(e === "TEXTJOIN" ? 2 : 0)
        .flatMap(L0)
        .filter((y) => !v || (y != null && y !== ""))
        .map(tn);
    return (
      g.reduce((y, b) => y + b.length, 0) +
        Math.max(0, g.length - 1) * h.length >
        32767 && o("#LIMIT!"),
      g.join(h)
    );
  }
  if (e === "NUMBERVALUE") {
    let h = [...tn(l[1] === undefined ? "." : l[1])][0],
      v = [...tn(l[2] === undefined ? "," : l[2])][0];
    (!h || !v || h === v || /[0-9+\-eE%]/.test(h) || /[0-9+\-eE%]/.test(v)) &&
      o();
    let g = m.replace(/ /g, "");
    if (!g) return 0;
    let y = /%+$/.exec(g)?.[0].length ?? 0;
    g = g.slice(0, g.length - y);
    let b = g.split(h);
    ((b.length > 2 || b[1]?.includes(v)) && o(),
      (g = b[0].split(v).join("") + (b.length === 2 ? "." + b[1] : "")),
      /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(g) || o());
    let x = Number(g);
    Number.isFinite(x) || o("#NUM!");
    for (let k = 0; k < y; k++) x /= 100;
    return x;
  }
  if (e === "VALUE") {
    let h = r.temporal ? r.temporal(m.trim()) : pO(m.trim());
    if (h !== undefined) return h;
    let v = m.trim(),
      g = false;
    v.startsWith("(") &&
      v.endsWith(")") &&
      ((g = true), (v = v.slice(1, -1).trim()));
    let y = v.endsWith("%");
    (y && (v = v.slice(0, -1).trim()),
      (v = v.replace(/^\$\s*/, "")),
      g && /^[+-]/.test(v) && o(),
      /^[+-]?(?:(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(
        v,
      ) || o());
    let b = (Number(v.replace(/,/g, "")) * (g ? -1 : 1)) / (y ? 100 : 1);
    return Number.isFinite(b) ? b : o("#NUM!");
  }
  if (e === "TEXT") {
    ((typeof l[1] != "string" || l[1].length > 256) && o(), f(l[0]));
    let h = r.format();
    return typeof h == "string" ? u(h) : h;
  }
  if (e === "TEXTBEFORE" || e === "TEXTAFTER") {
    let h = tn(l[1]),
      v = d(l[2], 1),
      g = f(l[3], 0),
      y = f(l[4], 0);
    if (((!v || ![0, 1].includes(g) || ![0, 1].includes(y)) && o(), !m))
      return "";
    if ((Math.abs(v) > [...m].length && o(), !h))
      return (e === "TEXTBEFORE") == v > 0 ? "" : m;
    let b = h.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"),
      x = new RegExp(v < 0 ? `(?=${b})` : b, g ? "giu" : "gu"),
      k = [...m.matchAll(x)].map((A) => A.index);
    if (v < 0) {
      let A = [],
        I = m.length;
      for (let N = k.length - 1; N >= 0; N--)
        k[N] + h.length <= I && (A.push(k[N]), (I = k[N]));
      k = A.reverse();
    }
    let w = k[v > 0 ? v - 1 : k.length + v];
    return w === undefined
      ? y && Math.abs(v) === k.length + 1
        ? (e === "TEXTBEFORE") == v > 0
          ? m
          : ""
        : l[5] !== undefined
          ? l[5]
          : o("#N/A")
      : e === "TEXTBEFORE"
        ? m.slice(0, w)
        : m.slice(w + h.length);
  }
  return Tu;
}
var IO = {
  maxIntervalsPerAxis: 1e4,
  maxFrozenVisibleRows: 100,
  maxFrozenVisibleCols: 50,
  maxXlsxHiddenRows: 1e5,
};
function kO(e) {
  if (!Number.isSafeInteger(e) || e < 0)
    throw Error("Axis capacity must be a nonnegative integer.");
}
function Vo(e, t) {
  if ((kO(t), e === undefined)) return [];
  if (!Array.isArray(e) || e.length > IO.maxIntervalsPerAxis)
    throw Error("Hidden indices require at most 10,000 intervals per axis.");
  let r = e
      .map((n) => {
        if (
          !n ||
          typeof n != "object" ||
          Array.isArray(n) ||
          (Object.getPrototypeOf(n) !== Object.prototype &&
            Object.getPrototypeOf(n) !== null) ||
          Object.keys(n).some((i) => i !== "start" && i !== "end") ||
          !Number.isInteger(n.start) ||
          !Number.isInteger(n.end) ||
          n.start < 0 ||
          n.end < n.start ||
          n.end >= t
        )
          throw Error(
            "Hidden intervals require valid zero-based inclusive axis bounds.",
          );
        return { start: n.start, end: n.end };
      })
      .sort((n, i) => n.start - i.start || n.end - i.end),
    o = [];
  for (let n of r) {
    let i = o.at(-1);
    i && n.start <= i.end + 1 ? (i.end = Math.max(i.end, n.end)) : o.push(n);
  }
  return o;
}
function Ss(e, t = []) {
  let r = Vo(t, e),
    o = [0];
  for (let c of r) o.push(o.at(-1) + c.end - c.start + 1);
  let n = e - o.at(-1),
    i = (c) => {
      let u = 0,
        f = r.length;
      for (; u < f;) {
        let d = (u + f) >>> 1;
        r[d].start <= c ? (u = d + 1) : (f = d);
      }
      return u - 1;
    },
    a = (c) => {
      let u = Math.max(0, Math.min(e, Math.floor(c))),
        f = i(u - 1),
        d =
          f < 0
            ? 0
            : o[f] + Math.min(u - r[f].start, r[f].end - r[f].start + 1);
      return u - d;
    },
    l = (c) => {
      let u = i(c);
      return u >= 0 && c <= r[u].end;
    },
    s = (c) => {
      if (!Number.isInteger(c) || c < 0 || c >= n) return -1;
      let u = 0,
        f = e - 1;
      for (; u < f;) {
        let d = Math.floor((u + f) / 2);
        a(d + 1) <= c ? (u = d + 1) : (f = d);
      }
      return u;
    };
  return {
    count: n,
    logicalCount: e,
    logicalIndex: s,
    countBefore: a,
    isHidden: l,
    viewIndex: (c) =>
      Number.isInteger(c) && c >= 0 && c < e && !l(c) ? a(c) : null,
    next: (c, u, f = false) => {
      let d = Math.trunc(c) + (f ? 0 : u);
      if (!n || (u === 1 && d >= e) || (u === -1 && d < 0)) return null;
      let p = u === 1 ? a(Math.max(0, d)) : a(Math.min(e, d + 1)) - 1,
        m = s(p);
      return m < 0 ? null : m;
    },
  };
}
var NO = (e, t) => (t === "row" ? e.rowCount : e.colCount);
function RO(e) {
  if (e !== "row" && e !== "column") throw Error("Choose rows or columns.");
}
function ea(e, t) {
  RO(t);
  let r = NO(e, t),
    o = [
      ...Vo(t === "row" ? e.hiddenRows : e.hiddenCols, r),
      ...(e.outlineGroups ?? [])
        .filter((i) => i.axis === t && i.collapsed)
        .map(({ start, end }) => ({ start, end })),
    ].sort((i, a) => i.start - a.start || i.end - a.end),
    n = [];
  for (let i of o) {
    let a = n.at(-1);
    a && i.start <= a.end + 1
      ? (a.end = Math.max(a.end, i.end))
      : n.push({ ...i });
  }
  return Vo(n, r);
}
var F0 = /[A-Za-z0-9_.\\\u0080-\uffff]/,
  CO =
    /^(?:'(?:[^']|'')*'|[A-Za-z_\\\u0080-\uffff][A-Za-z0-9_.\\\u0080-\uffff]*)(?::(?:'(?:[^']|'')*'|[A-Za-z_\\\u0080-\uffff][A-Za-z0-9_.\\\u0080-\uffff]*))?!/,
  SO = (e) =>
    [...e.toUpperCase()].reduce((t, r) => t * 26 + r.charCodeAt(0) - 64, 0) - 1;
function z0(e, t, r) {
  let o = SO(t) + (e ? 0 : r);
  return o < 0 || o >= yt.maxCols ? "#REF!" : e + Me(o);
}
function W0(e, t, r) {
  let o = Number(t) + (e ? 0 : r);
  return o < 1 || o > yt.maxRows ? "#REF!" : e + o;
}
function jo(e, t, r) {
  if (!Number.isSafeInteger(t) || !Number.isSafeInteger(r))
    throw new Error("Formula offsets must be integers.");
  if (!e.startsWith("=")) return e;
  let o = "",
    n = 0;
  for (; n < e.length;) {
    let i = e.slice(n),
      a = e[n],
      l = (n === 0 || !F0.test(e[n - 1])) && i.match(CO);
    if (l) {
      ((o += l[0]), (n += l[0].length));
      continue;
    }
    if (a === '"' || a === "'") {
      let p = n++;
      for (; n < e.length;)
        if (e[n++] === a)
          if (e[n] === a) n++;
          else break;
      o += e.slice(p, n);
      continue;
    }
    if (a === "[") {
      let p = n++,
        m = 1;
      for (; n < e.length && m;)
        (e[n] === "[" && m++, e[n] === "]" && m--, n++);
      o += e.slice(p, n);
      continue;
    }
    let s = n === 0 || !F0.test(e[n - 1]),
      c = s ? /^(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})/.exec(i) : null,
      u = s ? /^(\$?)([1-9]\d*):(\$?)([1-9]\d*)/.exec(i) : null,
      f = s ? /^(\$?)([A-Za-z]{1,3})(\$?)([1-9]\d*)/.exec(i) : null,
      d = c ?? u ?? f;
    if (
      d &&
      !F0.test(i[d[0].length] ?? "") &&
      !/^\s*[(!]/.test(i.slice(d[0].length))
    ) {
      if (c) o += `${z0(c[1], c[2], r)}:${z0(c[3], c[4], r)}`;
      else if (u) o += `${W0(u[1], u[2], t)}:${W0(u[3], u[4], t)}`;
      else if (f) {
        let p = z0(f[1], f[2], r),
          m = W0(f[3], f[4], t);
        o += p === "#REF!" || m === "#REF!" ? "#REF!" : p + m;
      }
      n += d[0].length;
      continue;
    }
    ((o += a), n++);
  }
  return o;
}
var S9 = Object.freeze({
  rulesPerSheet: 100,
  listItems: 200,
  checks: 1e5,
  customChecks: 5e3,
  displayedIssues: 500,
});
var $O = ["en-US", "en-GB", "de-DE", "fr-FR", "ja-JP", "en-IN"],
  OO = ["USD", "EUR", "GBP", "JPY", "CAD", "AUD", "INR", "CHF"],
  MO = [
    "general",
    "number",
    "currency",
    "accounting",
    "percent",
    "date",
    "time",
    "datetime",
    "duration",
    "custom",
  ],
  PO = {
    date: ["yyyy-mm-dd", "dd/mm/yyyy", "m/d/yyyy", "d mmm yyyy", "d mmmm yyyy"],
    time: ["hh:mm", "hh:mm:ss", "h:mm AM/PM", "hh:mm:ss.000"],
    datetime: [
      "yyyy-mm-dd hh:mm:ss",
      "dd/mm/yyyy hh:mm",
      "m/d/yyyy h:mm AM/PM",
    ],
    duration: ["[h]:mm", "[h]:mm:ss", "[h]:mm:ss.000", "[m]:ss", "[s].000"],
  };
function Ze(e) {
  throw new Error(e);
}
var UO = new Set(["number", "currency", "accounting", "percent"]),
  LO = new Set(["date", "time", "datetime", "duration"]);
function DO(e) {
  return (
    !!e &&
    typeof e == "object" &&
    !Array.isArray(e) &&
    [Object.prototype, null].includes(Object.getPrototypeOf(e))
  );
}
function mx(e) {
  (!DO(e) ||
    Object.keys(e).some(
      (i) =>
        ![
          "category",
          "locale",
          "currency",
          "decimals",
          "useGrouping",
          "scale",
          "negative",
          "pattern",
        ].includes(i),
    )) &&
    Ze("Number format must be a supported plain settings object.");
  let t = e;
  MO.includes(t.category) || Ze("Choose a supported number format category.");
  let r = t.category,
    o = t.locale ?? "en-US";
  $O.includes(o) || Ze("Choose one of the supported display locales.");
  let n = { category: r, locale: o };
  if (UO.has(r)) {
    if (r === "currency" || r === "accounting") {
      let s = t.currency ?? "USD";
      (OO.includes(s) || Ze("Choose a supported currency."), (n.currency = s));
    } else
      t.currency !== undefined &&
        Ze("Currency applies only to Currency or Accounting formats.");
    let a = t.decimals ?? (n.currency === "JPY" ? 0 : 2);
    ((typeof a != "number" || !Number.isInteger(a) || a < 0 || a > 12) &&
      Ze("Choose 0\u201312 decimal places."),
      (n.decimals = a),
      t.useGrouping !== undefined &&
        typeof t.useGrouping != "boolean" &&
        Ze("The grouping option must be true or false."),
      (n.useGrouping = t.useGrouping ?? r !== "percent"));
    let l = t.scale ?? 1;
    ((![1, 1e3, 1e6].includes(l) || (r === "percent" && l !== 1)) &&
      Ze("Choose units, thousands or millions; percentages use units."),
      (n.scale = l),
      t.negative !== undefined &&
        (typeof t.negative != "string" ||
          !["minus", "parentheses"].includes(t.negative)) &&
        Ze("Choose minus signs or parentheses for negatives."),
      r === "accounting" &&
        t.negative !== undefined &&
        t.negative !== "parentheses" &&
        Ze("Accounting negatives use parentheses."),
      (n.negative =
        r === "accounting" ? "parentheses" : (t.negative ?? "minus")),
      t.pattern !== undefined &&
        Ze("Use the Custom category to enter a format pattern."));
  } else if (
    ([t.currency, t.decimals, t.useGrouping, t.scale, t.negative].some(
      (i) => i !== undefined,
    ) && Ze("These numeric settings do not apply to the chosen category."),
    LO.has(r) || r === "custom")
  ) {
    let i = r === "custom" ? "0.00" : PO[r][0],
      a = t.pattern ?? i;
    ((typeof a != "string" ||
      !a ||
      a.length > 200 ||
      /[\u0000-\u001f\u007f]/.test(a)) &&
      Ze(
        "Enter a format pattern of 1\u2013200 characters without control characters.",
      ),
      (n.pattern = a));
    let l = Os(n.pattern);
    if (r !== "custom") {
      l.kind !== "temporal" &&
        Ze("Choose a date, time or elapsed-time pattern for this category.");
      let s = l.category;
      s !== r &&
        Ze(
          `This pattern is a ${s} format. Choose the matching category or Custom.`,
        );
    }
  } else
    t.pattern !== undefined && Ze("General does not accept a custom pattern.");
  return n;
}
function FO(e) {
  let t = [""],
    r = false;
  for (let o = 0; o < e.length; o++) {
    let n = e[o];
    if (n === "\\" && o + 1 < e.length) {
      t[t.length - 1] += n + e[++o];
      continue;
    }
    (n === '"' && (r = !r),
      n === ";" && !r ? t.push("") : (t[t.length - 1] += n));
  }
  return r || t.length > 4 ? undefined : t;
}
function zO(e) {
  let t = {
      prefix: "",
      suffix: "",
      minDecimals: 0,
      maxDecimals: 0,
      minIntegerDigits: 1,
      grouping: false,
      percent: false,
      hasNumber: false,
    },
    r = (o) => {
      t.hasNumber ? (t.suffix += o) : (t.prefix += o);
    };
  for (let o = 0; o < e.length;) {
    let n = e[o];
    if (n === '"') {
      let i = e.indexOf('"', o + 1);
      if (i < 0) return;
      (r(e.slice(o + 1, i)), (o = i + 1));
    } else if (n === "\\") {
      if (o + 1 >= e.length) return;
      (r(e[o + 1]), (o += 2));
    } else if (n === "[") {
      let i =
        /^\[(?:Black|Blue|Cyan|Green|Magenta|Red|White|Yellow|Color\d+)\]/i.exec(
          e.slice(o),
        );
      if (!i) return;
      o += i[0].length;
    } else if (n === "0" || n === "#") {
      if (t.hasNumber) return;
      let i = /^[0#.,]+/.exec(e.slice(o))[0],
        a = /^(0+|#+0*|#,##0+)(?:\.(0*#*))?$/.exec(i);
      if (!a) return;
      let l = a[2] ?? "";
      if (
        ((t.minDecimals = [...l].filter((s) => s === "0").length),
        (t.maxDecimals = l.length),
        (t.minIntegerDigits = [...a[1]].filter((s) => s === "0").length),
        t.maxDecimals > 20 || t.minIntegerDigits > 21)
      )
        return;
      ((t.grouping = a[1].includes(",")),
        (t.hasNumber = true),
        (o += i.length));
    } else if (n === "%") {
      if (t.percent) return;
      ((t.percent = true), r(n), o++);
    } else if (/[$€£¥ ()+\-]/.test(n)) (r(n), o++);
    else return;
  }
  return t;
}
var $s = new Map();
function WO(e) {
  let t = e.replace(/"[^"]*"|\\./g, "");
  if (!/[ymdhs]/i.test(t)) return;
  let r = [],
    o = 0;
  for (let s = 0; s < e.length;) {
    let c = e.slice(s),
      u = c[0];
    if (u === '"') {
      let h = e.indexOf('"', s + 1);
      (h < 0 && Ze("The format contains an unclosed quoted literal."),
        r.push({ kind: "literal", text: e.slice(s + 1, h) }),
        (s = h + 1));
      continue;
    }
    if (u === "\\") {
      (c.length < 2 && Ze("A format escape must be followed by a character."),
        r.push({ kind: "literal", text: c[1] }),
        (s += 2));
      continue;
    }
    let f = /^(AM\/PM|A\/P)/i.exec(c);
    if (f) {
      (r.push({ kind: "ampm", text: f[0] }), (s += f[0].length));
      continue;
    }
    let d = /^\[(h{1,2}|m{1,2}|s{1,2})\]/i.exec(c);
    if (d) {
      (r.push({ kind: "elapsed", text: d[1].toLowerCase() }),
        (s += d[0].length));
      continue;
    }
    let p = /^(y+|m+|d+|h+|s+)/i.exec(c);
    if (p) {
      let h = p[0].toLowerCase(),
        v = h[0];
      (((v === "y" && ![2, 4].includes(h.length)) ||
        (v === "m" && h.length > 4) ||
        (["d", "h", "s"].includes(v) && h.length > 2)) &&
        Ze(
          "This date pattern uses unsupported fields. Use yy/yyyy, m\u2013mmmm, d/dd, h/hh or s/ss.",
        ),
        r.push({
          kind: { y: "year", m: "month", d: "day", h: "hour", s: "second" }[v],
          text: h,
        }),
        (s += h.length));
      continue;
    }
    let m = /^\.(0{1,3})(?!0)/.exec(c);
    if (m) {
      ((!r.length ||
        !["second", "elapsed"].includes(r.at(-1).kind) ||
        (r.at(-1).kind === "elapsed" && !r.at(-1).text.startsWith("s")) ||
        o) &&
        Ze(
          "Fractional seconds must follow seconds and use 1\u20133 zero digits.",
        ),
        (o = m[1].length),
        r.push({ kind: "fraction", text: m[1] }),
        (s += m[0].length));
      continue;
    }
    if (/^[ /:,.-]$/.test(u)) {
      (r.push({ kind: "literal", text: u }), s++);
      continue;
    }
    Ze(
      "Unsupported date/time code. Quote literal words and use one supported date/time pattern.",
    );
  }
  let n = r.filter((s) => s.kind !== "literal");
  n.forEach((s, c) => {
    s.kind === "month" &&
      s.text.length <= 2 &&
      (n[c - 1]?.kind === "hour" ||
        (n[c - 1]?.kind === "elapsed" && n[c - 1].text.startsWith("h")) ||
        n[c + 1]?.kind === "second") &&
      (s.kind = "minute");
  });
  let i = n.some((s) => ["year", "month", "day"].includes(s.kind)),
    a = n.some((s) => ["hour", "minute", "second"].includes(s.kind)),
    l = n.filter((s) => s.kind === "elapsed");
  if (l.length) {
    (l.length !== 1 ||
      i ||
      n.some((c) => ["hour", "ampm"].includes(c.kind)) ||
      n[0].kind !== "elapsed") &&
      Ze(
        "Elapsed-time patterns start with one [h], [m] or [s] field and cannot include calendar or AM/PM fields.",
      );
    let s = l[0].text[0];
    ((s === "s" && a) || (s === "m" && n.some((c) => c.kind === "minute"))) &&
      Ze("Use [h]:mm:ss, [m]:ss or [s] for elapsed time.");
  }
  return (
    n.some((s) => s.kind === "ampm") &&
      !n.some((s) => s.kind === "hour") &&
      Ze("AM/PM requires an hour field."),
    !i &&
      !a &&
      !l.length &&
      Ze("The format does not contain a date or time field."),
    {
      kind: "temporal",
      category: l.length
        ? "duration"
        : i && a
          ? "datetime"
          : i
            ? "date"
            : "time",
      tokens: r,
      fraction: o,
    }
  );
}
function BO(e) {
  let t = "",
    r = 1,
    o = false;
  for (let i = 0; i < e.length;) {
    let a = e.slice(i);
    if (a[0] === '"') {
      let l = e.indexOf('"', i + 1);
      (l < 0 && Ze("The format contains an unclosed quoted literal."),
        (t += e.slice(i, l + 1)),
        (i = l + 1));
    } else if (a[0] === "\\")
      (a.length < 2 && Ze("A format escape must be followed by a character."),
        (t += a.slice(0, 2)),
        (i += 2));
    else if (/^[0#]/.test(a)) {
      o && Ze("Only one numeric field per custom section is supported.");
      let l = /^[0#.,]+/.exec(a)[0],
        s = /,+$/.exec(l)?.[0] ?? "";
      (s.length > 2 &&
        Ze("Custom display scaling supports at most two trailing commas."),
        (r = 1e3 ** s.length),
        (t += l.slice(0, l.length - s.length)),
        (i += l.length),
        (o = true));
    } else ((t += a[0]), i++);
  }
  /\[/.test(t.replace(/"[^"]*"|\\./g, "")) &&
    Ze(
      "Custom colors, conditions and locale tags are not supported. Use conditional formatting or the locale selector.",
    );
  let n = zO(t);
  return (
    n ||
      Ze(
        "Unsupported custom format. Use 0/# digits, grouping, decimals, percent, quoted literals and positive/negative/zero sections.",
      ),
    n.maxDecimals > 12 &&
      Ze("Custom formats support at most 12 decimal places."),
    { parsed: n, divisor: r }
  );
}
function Os(e) {
  let t = $s.get(e);
  if (t) return t;
  (typeof e != "string" ||
    !e.length ||
    e.length > 200 ||
    /[\u0000-\u001f\u007f]/.test(e)) &&
    Ze("Enter a supported format pattern of 1\u2013200 characters.");
  let r,
    o = FO(e);
  o || Ze("A format has at most four sections and must close quoted literals.");
  let n = o.length === 1 ? WO(e) : undefined;
  return (
    n
      ? (r = n)
      : (o.length === 4 &&
          !VO(o[3]) &&
          Ze(
            "The text section supports @ and quoted or escaped literals only.",
          ),
        (r = {
          kind: "numeric",
          sections: o.slice(0, 3).map(BO),
          ...(o.length === 4 ? { textSection: o[3] } : {}),
        })),
    $s.size >= 256 && $s.delete($s.keys().next().value),
    $s.set(e, r),
    r
  );
}
function VO(e) {
  return e.replace(/"[^"]*"|\\./g, "").replace(/[@ ]/g, "") === "";
}
function jO(e, t) {
  return e.replace(
    /"([^"]*)"|\\(.)|@/g,
    (r, o, n) => o ?? n ?? (r === "@" ? t : r),
  );
}
var px = (e) =>
  new Intl.NumberFormat(e).formatToParts(1.1).find((t) => t.type === "decimal")
    ?.value ?? ".";
function ZO(e, t, r) {
  let o = e < 0 && t.sections.length > 1,
    { parsed, divisor } =
      t.sections[e === 0 && t.sections.length > 2 ? 2 : o ? 1 : 0];
  if (!parsed.hasNumber) return parsed.prefix + parsed.suffix;
  let a = (Math.abs(e) / divisor) * (parsed.percent ? 100 : 1);
  if (!Number.isFinite(a)) return "#NUM!";
  let l = new Intl.NumberFormat(r, {
    minimumFractionDigits: parsed.minDecimals,
    maximumFractionDigits: parsed.maxDecimals,
    minimumIntegerDigits: Math.max(1, parsed.minIntegerDigits),
    useGrouping: parsed.grouping,
  }).format(a);
  return (
    parsed.minIntegerDigits ||
      (l === "0" && !parsed.minDecimals
        ? (l = "")
        : l.startsWith(`0${px(r)}`) && (l = l.slice(1))),
    (e < 0 && !o ? "-" : "") + parsed.prefix + l + parsed.suffix
  );
}
function GO(e, t, r, o) {
  let n = t.category === "duration",
    i = 2958465 - (o === "1904" && !n ? 1462 : 0);
  if (Math.abs(e) >= i + 1 || (!n && e < 0)) return "#NUM!";
  let a = t.category !== "date",
    l = 10 ** t.fraction,
    s = t.tokens.some(
      (R) =>
        R.kind === "second" || (R.kind === "elapsed" && R.text.startsWith("s")),
    ),
    c = a
      ? s
        ? Math.round(Math.abs(e) * 86400 * l) / l
        : Math.floor(Math.abs(e) * 86400)
      : Math.floor(e) * 86400,
    u = Math.floor(c / 86400),
    f = c - u * 86400;
  if (u > i) return "#NUM!";
  let d = Math.round(f * 1e3),
    p = Math.floor(d / 36e5),
    m = Math.floor(d / 6e4) % 60,
    h = Math.floor(d / 1e3) % 60,
    v = Math.floor((d % 1e3) / 10 ** (3 - t.fraction)),
    g = o === "1900" && u === 60,
    y =
      o === "1904"
        ? Date.UTC(1904, 0, 1) + u * 864e5
        : Date.UTC(1899, 11, 31) + (u > 60 ? u - 1 : u) * 864e5,
    b = new Date(y),
    x = g ? 1900 : b.getUTCFullYear(),
    k = g ? 1 : b.getUTCMonth(),
    w = g ? 29 : b.getUTCDate(),
    A = t.tokens.some((R) => R.kind === "ampm"),
    I = (R, C) => String(R).padStart(C, "0"),
    N = t.tokens
      .map((R) => {
        switch (R.kind) {
          case "literal":
            return R.text;
          case "year":
            return R.text.length === 2 ? I(x % 100, 2) : I(x, 4);
          case "month":
            return R.text.length < 3
              ? I(k + 1, R.text.length)
              : new Intl.DateTimeFormat(r, {
                  month: R.text.length === 3 ? "short" : "long",
                  timeZone: "UTC",
                }).format(new Date(Date.UTC(2e3, k, 1)));
          case "day":
            return I(w, R.text.length);
          case "hour":
            return I(A ? p % 12 || 12 : p, R.text.length);
          case "minute":
            return I(m, R.text.length);
          case "second":
            return I(h, R.text.length);
          case "fraction":
            return px(r) + I(v, t.fraction);
          case "ampm": {
            let C = p < 12 ? "AM" : "PM",
              T = R.text.length === 3 ? C[0] : C;
            return R.text === R.text.toLowerCase() ? T.toLowerCase() : T;
          }
          case "elapsed":
            return I(
              Math.floor(
                c / (R.text[0] === "h" ? 3600 : R.text[0] === "m" ? 60 : 1),
              ),
              R.text.length,
            );
        }
      })
      .join("");
  return (n && e < 0 ? "-" : "") + N;
}
function hx(e, t, r = "1900") {
  let o = mx(t);
  if (e === null) return { text: "" };
  if (typeof e == "boolean") return { text: e ? "TRUE" : "FALSE" };
  if (typeof e == "string") {
    if (o.category === "custom" && !/^#[A-Z0-9/?!]+$/.test(e)) {
      let c = Os(o.pattern);
      if (c.kind === "numeric" && c.textSection !== undefined)
        return { text: jO(c.textSection, e) };
    }
    return { text: e };
  }
  if (!Number.isFinite(e)) return { text: "#NUM!" };
  if (o.category === "general")
    return {
      text: new Intl.NumberFormat(o.locale, {
        maximumSignificantDigits: 15,
        useGrouping: false,
      }).format(e === 0 ? 0 : e),
    };
  if (o.pattern) {
    let c = Os(o.pattern);
    return {
      text: c.kind === "temporal" ? GO(e, c, o.locale, r) : ZO(e, c, o.locale),
    };
  }
  let n = e / (o.scale ?? 1),
    i = o.scale === 1e3 ? " K" : o.scale === 1e6 ? " M" : "";
  if (o.category === "percent" && !Number.isFinite(n * 100))
    return { text: "#NUM!" };
  let a = {
    minimumFractionDigits: o.decimals,
    maximumFractionDigits: o.decimals,
    useGrouping: o.useGrouping,
  };
  if (o.category === "accounting") {
    let u = new Intl.NumberFormat(o.locale, {
        ...a,
        style: "currency",
        currency: o.currency,
      }).formatToParts(Math.abs(n)),
      f = u.find((h) => h.type === "currency")?.value ?? o.currency,
      d =
        u.findIndex((h) => h.type === "currency") <
        u.findIndex((h) => h.type === "integer")
          ? "start"
          : "end",
      p = new Intl.NumberFormat(o.locale, a).format(Math.abs(n)),
      m = (n === 0 ? "-" : n < 0 ? `(${p})` : p) + i;
    return {
      text: d === "start" ? `${f} ${m}` : `${m} ${f}`,
      accounting: { symbol: f, amount: m, symbolPosition: d },
    };
  }
  let l = n < 0 && o.negative === "parentheses",
    s = new Intl.NumberFormat(o.locale, {
      ...a,
      ...(o.category === "currency"
        ? { style: "currency", currency: o.currency }
        : o.category === "percent"
          ? { style: "percent" }
          : {}),
    }).format(l ? Math.abs(n) : n);
  return { text: (l ? `(${s})` : s) + i };
}
var $y = rT(qN(), 1);
function HN(e) {
  let [t, r = "0"] = Math.abs(e).toString().split("e"),
    [o, n = ""] = t.split(".");
  return {
    coefficient: BigInt(o + n) * (e < 0 ? -1n : 1n),
    exponent: Number(r) - n.length,
  };
}
function KN(e) {
  let t = new Map();
  for (let i of e) {
    if (!Number.isFinite(i)) return;
    let { coefficient, exponent } = HN(i);
    t.set(exponent, (t.get(exponent) ?? 0n) + coefficient);
  }
  let r = [...t].filter(([, i]) => i !== 0n);
  if (!r.length) return { coefficient: 0n, exponent: 0 };
  let o = r.reduce((i, [a]) => Math.min(i, a), 1 / 0);
  return {
    coefficient: r.reduce((i, [a, l]) => i + l * 10n ** BigInt(a - o), 0n),
    exponent: o,
  };
}
function Nr(e) {
  let t = KN(e);
  return t ? +`${t.coefficient}e${t.exponent}` : NaN;
}
function Zt(e) {
  if (!e.length) return NaN;
  let t = KN(e);
  if (!t) return NaN;
  let r = t.coefficient < 0n ? "-" : "",
    o = t.coefficient < 0n ? -t.coefficient : t.coefficient,
    n = BigInt(e.length),
    i = o / n,
    a = (((o % n) * 10n ** 40n) / n).toString().padStart(40, "0");
  return +`${r}${i}.${a}e${t.exponent}`;
}
function _f(e) {
  if (e.some((o) => !Number.isFinite(o))) return NaN;
  if (e.some((o) => o === 0)) return 0;
  let t = 1n,
    r = 0;
  for (let o of e) {
    let n = HN(o);
    if (((t *= n.coefficient), (r += n.exponent), e.length > 255)) {
      let i = t < 0n,
        a = (i ? -t : t).toString();
      a.length > 80 &&
        ((r += a.length - 80), (t = BigInt(a.slice(0, 80)) * (i ? -1n : 1n)));
    }
  }
  return +`${t}e${r}`;
}
var cy = Symbol("not-arithmetic"),
  x4 = new Set([
    "ROUND",
    "ROUNDUP",
    "ROUNDDOWN",
    "TRUNC",
    "MROUND",
    "CEILING",
    "FLOOR",
    "SUMPRODUCT",
  ]);
function ly(e) {
  let [t, r = "0"] = Math.abs(e).toString().split("e"),
    [o, n = ""] = t.split(".");
  return { coefficient: BigInt(o + n), exponent: Number(r) - n.length };
}
function XN(e, t, r) {
  let o = e / t,
    n = e % t;
  return (
    o +
    (n !== 0n && (r === "away" || (r === "nearest" && n * 2n >= t)) ? 1n : 0n)
  );
}
function I4(e, t, r) {
  if (((t = Math.trunc(t)), e === 0 || t >= 324)) return e;
  if (t <= -309) return r === "away" ? 1 / 0 : 0;
  let o = ly(e),
    n = o.exponent + t;
  if (n >= 0) return e;
  let i = XN(o.coefficient, 10n ** BigInt(-n), r);
  return Math.sign(e) * +`${i}e${-t}`;
}
function k4(e, t, r) {
  let o = ly(e),
    n = ly(t),
    i = Math.min(o.exponent, n.exponent),
    a = o.coefficient * 10n ** BigInt(o.exponent - i),
    l = n.coefficient * 10n ** BigInt(n.exponent - i),
    s = XN(a, l, r);
  return Math.sign(e) * +`${s * n.coefficient}e${n.exponent}`;
}
function JN(e, t, r) {
  if (!x4.has(e)) return cy;
  let { retrieve, issue, isIssue } = r,
    a = t.map(retrieve),
    l = (d) => (Number.isFinite(d) ? (d === 0 ? 0 : d) : issue("#NUM!"));
  if (e === "SUMPRODUCT") {
    if (!t.length || t.length > 255) return issue("#VALUE!");
    let d = a.map((v) =>
        Array.isArray(v) ? (v.every(Array.isArray) ? v : [v]) : [[v]],
      ),
      p = d[0].length,
      m = d[0][0]?.length ?? 0;
    if (
      !p ||
      !m ||
      d.some((v) => v.length !== p || v.some((g) => g.length !== m))
    )
      return issue("#VALUE!");
    if (p * m * d.length > 1e5) return issue("#LIMIT!");
    let h = [];
    for (let v = 0; v < p; v++)
      for (let g = 0; g < m; g++) {
        let y = [];
        for (let x of d) {
          let k = x[v][g];
          if (isIssue(k)) return k;
          y.push(typeof k == "number" ? k : 0);
        }
        let b = _f(y);
        if (!Number.isFinite(b)) return issue("#NUM!");
        h.push(b);
      }
    return l(Nr(h));
  }
  if (t.length < (e === "TRUNC" ? 1 : 2) || t.length > 2)
    return issue("#VALUE!");
  let s = [];
  for (let d of a) {
    if (isIssue(d)) return d;
    if (
      Array.isArray(d) ||
      (typeof d == "string" &&
        !/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?$/i.test(d.trim()))
    )
      return issue("#VALUE!");
    let p = d == null ? 0 : Number(d);
    if (!Number.isFinite(p)) return issue("#NUM!");
    s.push(p);
  }
  let [c, u = 0] = s;
  if (["ROUND", "ROUNDUP", "ROUNDDOWN", "TRUNC"].includes(e))
    return l(
      I4(c, u, e === "ROUND" ? "nearest" : e === "ROUNDUP" ? "away" : "toward"),
    );
  if (c === 0 || u === 0) return 0;
  if (
    (e === "MROUND" && Math.sign(c) !== Math.sign(u)) ||
    (e === "FLOOR" && c > 0 && u < 0)
  )
    return issue("#NUM!");
  let f =
    e === "MROUND"
      ? "nearest"
      : e === "CEILING"
        ? c < 0 && u > 0
          ? "toward"
          : "away"
        : c < 0 && u > 0
          ? "away"
          : "toward";
  return l(k4(c, u, f));
}
var uy = ["LINEST", "LOGEST", "TREND", "GROWTH"],
  fy = Symbol("not-regression"),
  rl = {
    values: 4e4,
    observations: 1e4,
    predictors: 32,
    work: 4e6,
    output: 1e4,
  },
  Hr = (e, t) => {
    let r = 0,
      o = 0;
    for (let n = 0; n < e.length; n++) {
      let i = e[n] * t[n],
        a = r + i;
      ((o += Math.abs(r) >= Math.abs(i) ? r - a + i : i - a + r), (r = a));
    }
    return r + o;
  };
function YN(e, t, r) {
  if (!uy.includes(e)) return fy;
  let o = (z = "#VALUE!") => {
      throw r.issue(z);
    },
    n = (z) => (Number.isFinite(z) ? z : o("#NUM!"));
  (!t.length || t.length > 4 || t[0] == null) && o();
  let i = 0,
    a = (z) => {
      if (r.isIssue(z)) throw z;
      let J = Array.isArray(z) ? (z.every(Array.isArray) ? z : [z]) : [[z]],
        ne = J[0]?.length ?? 0;
      return (
        (!J.length || !ne || J.some((ce) => ce.length !== ne)) && o(),
        (i += J.length * ne),
        i > rl.values && o("#LIMIT!"),
        J.map((ce) =>
          ce.map((Te) => {
            if (r.isIssue(Te)) throw Te;
            return (typeof Te != "number" && o(), n(Te));
          }),
        )
      );
    },
    l = (z, J) => {
      if (z == null) return J;
      let ne = r.retrieve(z);
      if (r.isIssue(ne)) throw ne;
      return ne == null
        ? false
        : typeof ne == "boolean"
          ? ne
          : typeof ne == "number"
            ? n(ne) !== 0
            : typeof ne == "string" && /^(true|false)$/i.test(ne)
              ? /^true$/i.test(ne)
              : o();
    },
    s = e === "TREND" || e === "GROWTH",
    c = e === "LOGEST" || e === "GROWTH",
    u = l(t[s ? 3 : 2], true),
    f = !s && l(t[3], false),
    d = a(r.retrieve(t[0])),
    p = d.length * d[0].length;
  p > rl.observations && o("#LIMIT!");
  let m = d.flat();
  c && m.some((z) => z <= 0) && o("#NUM!");
  let h = c ? m.map(Math.log) : m,
    v =
      t[1] == null
        ? d.map((z, J) => z.map((ne, ce) => J * z.length + ce + 1))
        : a(r.retrieve(t[1])),
    g = v.length === d.length && v[0].length === d[0].length,
    y = d.length === 1,
    b;
  if (g) b = v.flat().map((z) => [z]);
  else if (d[0].length === 1 && v.length === p) b = v;
  else if (y && v[0].length === p)
    b = Array.from({ length: p }, (z, J) => v.map((ne) => ne[J]));
  else return o("#REF!");
  let x = b[0].length;
  (x > rl.predictors || p * x * x > rl.work) && o("#LIMIT!");
  let k = Array.from({ length: x }, (z, J) =>
      u ? Zt(b.map((ne) => ne[J])) : 0,
    ),
    w = u ? Zt(h) : 0,
    A = h.map((z) => n(z - w)),
    I = Math.max(...A.map(Math.abs)) || 1,
    N = A.map((z) => z / I),
    R = [],
    C = Array.from({ length: x }, (z, J) => {
      let ne = b.map((Te) => n(Te[J] - k[J])),
        ce = Math.max(...ne.map(Math.abs)) || 1;
      return (R.push(ce), ne.map((Te) => Te / ce));
    }),
    T = C.map((z) => [...z]),
    _ = Array.from({ length: x }, (z, J) => J),
    P = Array.from({ length: x }, () => Array(x).fill(0)),
    S = [],
    F = [],
    D = 1e-12 * Math.max(1, ...C.map((z) => Math.sqrt(Hr(z, z)))),
    Q = 0;
  for (let z = 0; z < x; z++) {
    let J = z;
    for (let ce = z + 1; ce < x; ce++)
      Hr(C[ce], C[ce]) > Hr(C[J], C[J]) && (J = ce);
    (([C[z], C[J]] = [C[J], C[z]]), ([_[z], _[J]] = [_[J], _[z]]));
    for (let ce = 0; ce < z; ce++) [P[ce][z], P[ce][J]] = [P[ce][J], P[ce][z]];
    let ne = Math.sqrt(Hr(C[z], C[z]));
    if (ne <= D) break;
    ((P[z][z] = ne),
      S.push(C[z].map((ce) => ce / ne)),
      F.push(Hr(S[z], N)),
      Q++);
    for (let ce = z + 1; ce < x; ce++)
      for (let Te = 0; Te < 2; Te++) {
        let rt = Hr(S[z], C[ce]);
        ((P[z][ce] += rt), (C[ce] = C[ce].map((Le, Ke) => Le - rt * S[z][Ke])));
      }
  }
  let te = (z) => {
      let J = Array(Q).fill(0);
      for (let ne = Q - 1; ne >= 0; ne--) {
        let ce = z[ne];
        for (let Te = ne + 1; Te < Q; Te++) ce -= P[ne][Te] * J[Te];
        J[ne] = n(ce / P[ne][ne]);
      }
      return J;
    },
    K = Array(x).fill(0);
  te(F).forEach((z, J) => (K[_[J]] = z));
  let Z = K.map((z, J) => n(z * (I / R[J]))),
    B = u ? n(w - Hr(Z, k)) : 0,
    X = (z) => {
      let J = n(Math.exp(z));
      return J > 0 ? J : o("#NUM!");
    };
  if (s) {
    let z = t[2] == null ? v : a(r.retrieve(t[2])),
      J,
      ne,
      ce;
    if (x === 1)
      ((J = z.flat().map((rt) => [rt])), (ne = z.length), (ce = z[0].length));
    else if (!y && z[0].length === x) ((J = z), (ne = J.length), (ce = 1));
    else if (y && z.length === x)
      ((J = Array.from({ length: z[0].length }, (rt, Le) =>
        z.map((Ke) => Ke[Le]),
      )),
        (ne = 1),
        (ce = J.length));
    else return o("#REF!");
    ne * ce > rl.output && o("#LIMIT!");
    let Te = J.map((rt) => {
      let Le = n(
        w +
          Hr(
            K,
            rt.map((Ke, Ct) => n(Ke - k[Ct]) / R[Ct]),
          ) *
            I,
      );
      return c ? X(Le) : Le;
    });
    return Array.from({ length: ne }, (rt, Le) =>
      Te.slice(Le * ce, (Le + 1) * ce),
    );
  }
  let L = [...Z]
    .reverse()
    .concat(B)
    .map((z) => (c ? X(z) : z));
  if (!f) return [L];
  let U = N.map(
      (z, J) =>
        z -
        Hr(
          K,
          T.map((ne) => ne[J]),
        ),
    ),
    Y = Hr(U, U),
    ie = Hr(N, N);
  Y <= 1e-26 * Math.max(ie, Number.MIN_VALUE) && (Y = 0);
  let ve = p - Q - Number(u),
    $ = ve > 0 ? Math.sqrt(Y / ve) : null,
    W = () => r.issue("#N/A"),
    se = () => r.issue("#DIV/0!"),
    me = Array.from({ length: Q }, (z, J) =>
      te(Array.from({ length: Q }, (ne, ce) => (ce === J ? 1 : 0))),
    ),
    we = Array(x).fill(0);
  for (let z = 0; z < Q; z++)
    we[_[z]] =
      $ === null
        ? se()
        : n(Math.hypot(...me.map((J) => J[z])) * $ * (I / R[_[z]]));
  let ue = _.slice(0, Q).map((z) => k[z] / R[z]),
    xe = 1 / p + me.reduce((z, J) => z + Hr(ue, J) ** 2, 0),
    re = [...we]
      .reverse()
      .concat(u ? ($ === null ? se() : n(Math.sqrt(xe) * $ * I)) : W()),
    Ee = Math.max(0, ie - Y),
    fe = (z) => n(z * I * I),
    ke = (z, J) => [z, J, ...Array.from({ length: x - 1 }, W)];
  return [
    L,
    re,
    ke(
      ie ? Math.max(0, Math.min(1, Ee / ie)) : se(),
      $ === null ? se() : n($ * I),
    ),
    ke(ve > 0 && Q && Y ? n(Ee / Q / (Y / ve)) : se(), ve),
    ke(fe(Ee), fe(Y)),
  ];
}
var dy = ["T.TEST", "Z.TEST"],
  my = Symbol("not-inference");
function QN(e, t, r) {
  if (!dy.includes(e)) return my;
  let o = (I = "#VALUE!") => {
      throw r.issue(I);
    },
    n = (I) => (Number.isFinite(I) ? I : o("#NUM!"));
  ((e === "T.TEST" && t.length !== 4) ||
    (e === "Z.TEST" && (t.length < 2 || t.length > 3))) &&
    o();
  let i = t.map(r.retrieve),
    a = 0,
    l = (I) => {
      if (Array.isArray(I)) return I.flatMap(l);
      if ((++a > 4e4 && o("#LIMIT!"), r.isIssue(I))) throw I;
      return (typeof I == "number" && n(I), [I]);
    },
    s = (I) => {
      if (r.isIssue(I)) throw I;
      if (Array.isArray(I) || I == null || (typeof I == "string" && !I.trim()))
        return o();
      let N = Number(I);
      return Number.isFinite(N) ? N : o();
    },
    c = (I) => I.filter((N) => typeof N == "number"),
    u = (I) => {
      if (I.length < 2) return o("#DIV/0!");
      let N = Zt(I);
      return {
        mean: N,
        variance: n(Nr(I.map((R) => (R - N) ** 2)) / (I.length - 1)),
      };
    },
    f = (I) => {
      if (r.isIssue(I)) throw I;
      return typeof I != "number" ||
        !Number.isFinite(I) ||
        I < -1e-12 ||
        I > 1 + 1e-12
        ? o("#NUM!")
        : Math.max(0, Math.min(1, I));
    },
    d = l(i[0]);
  if (e === "Z.TEST") {
    let I = c(d),
      N = I.length;
    if (!N) return o("#N/A");
    let R = s(i[1]),
      C = Math.max(Math.abs(R), ...I.map(Math.abs)) || 1,
      T = I.map((F) => F / C),
      _ = t[2] == null ? Math.sqrt(u(T).variance) : s(i[2]) / C;
    if (_ < 0) return o("#NUM!");
    if (!_) return o("#DIV/0!");
    let P = n(((Zt(T) - R / C) / _) * Math.sqrt(N)),
      S = r.distribution("ERFC", [P / Math.SQRT2]);
    if (r.isIssue(S)) throw S;
    return typeof S == "number" ? f(S / 2) : o("#NUM!");
  }
  let p = l(i[1]),
    m = Math.trunc(s(i[2])),
    h = Math.trunc(s(i[3]));
  ((![1, 2].includes(m) || ![1, 2, 3].includes(h)) && o("#NUM!"),
    h === 1 && d.length !== p.length && o("#N/A"));
  let v, g;
  if (h === 1) {
    let I = d
      .map((N, R) => [N, p[R]])
      .filter((N) => N.every((R) => typeof R == "number"));
    ((v = I.map((N) => N[0])), (g = I.map((N) => N[1])));
  } else ((v = c(d)), (g = c(p)));
  (v.length < 2 || g.length < 2) && o("#DIV/0!");
  let y = Math.max(...v.map(Math.abs), ...g.map(Math.abs)) || 1;
  ((v = v.map((I) => I / y)), (g = g.map((I) => I / y)));
  let b, x, k;
  if (h === 1) {
    let I = u(v.map((N, R) => N - g[R]));
    ((b = I.mean), (x = I.variance / v.length), (k = v.length - 1));
  } else {
    let I = u(v),
      N = u(g);
    if (((b = I.mean - N.mean), h === 2))
      ((k = v.length + g.length - 2),
        (x =
          (((v.length - 1) * I.variance + (g.length - 1) * N.variance) / k) *
          (1 / v.length + 1 / g.length)));
    else {
      let R = I.variance / v.length,
        C = N.variance / g.length;
      ((x = R + C),
        (k = x ** 2 / (R ** 2 / (v.length - 1) + C ** 2 / (g.length - 1))));
    }
  }
  x || o("#DIV/0!");
  let w = n(Math.abs(b) / Math.sqrt(x));
  n(k);
  let A = r.distribution("BETA.DIST", [k / (k + w * w), k / 2, 0.5, true]);
  return (f(A) * m) / 2;
}
var py = ["CONVERT", "SLOPE", "RSQ", "STEYX", "PEARSON"],
  hy = Symbol("not-specialist"),
  nl = [],
  ge = (e, t, r, o = false, n = 0, i = 1) => {
    for (let a of e.split("|"))
      nl.push({
        symbol: a,
        dimension: t,
        factor: r,
        metric: o,
        offset: n,
        power: i,
      });
  };
ge("g", "mass", 0.001, true);
ge("lbm", "mass", 0.45359237);
ge("ozm", "mass", 0.45359237 / 16);
ge("grain", "mass", 0.45359237 / 7e3);
ge("stone", "mass", 0.45359237 * 14);
ge("ton", "mass", 0.45359237 * 2e3);
ge("uk_ton|LTON|brton", "mass", 0.45359237 * 2240);
ge("cwt|shweight", "mass", 0.45359237 * 100);
ge("uk_cwt|lcwt|hweight", "mass", 0.45359237 * 112);
ge("m", "length", 1, true);
ge("in", "length", 0.0254);
ge("ft", "length", 0.3048);
ge("yd", "length", 0.9144);
ge("mi", "length", 1609.344);
ge("Nmi", "length", 1852);
ge("ang", "length", 1e-10, true);
ge("Picapt|Pica", "length", 0.0254 / 72);
ge("pica", "length", 0.0254 / 6);
ge("sec|s", "time", 1, true);
ge("mn|min", "time", 60);
ge("hr", "time", 3600);
ge("day|d", "time", 86400);
ge("yr", "time", 365.25 * 86400);
ge("Pa|p", "pressure", 1, true);
ge("atm|at", "pressure", 101325);
ge("psi", "pressure", 6894.757293168361);
ge("N", "force", 1, true);
ge("dyn|dy", "force", 1e-5, true);
ge("lbf", "force", 4.4482216152605);
ge("J", "energy", 1, true);
ge("e", "energy", 1e-7, true);
ge("c", "energy", 4.184, true);
ge("cal", "energy", 4.1868, true);
ge("Wh|wh", "energy", 3600, true);
ge("BTU|btu", "energy", 1055.05585262);
ge("W|w", "power", 1, true);
ge("HP|h", "power", 745.6998715822702);
ge("PS", "power", 735.49875);
ge("T", "magnetism", 1, true);
ge("ga", "magnetism", 1e-4, true);
ge("K|kel", "temperature", 1, true);
ge("C|cel", "temperature", 1, false, 273.15);
ge("F|fah", "temperature", 5 / 9, false, 273.15 - 160 / 9);
ge("Rank", "temperature", 5 / 9);
ge("Reau", "temperature", 1.25, false, 273.15);
ge("l|L|lt", "volume", 0.001, true);
ge("gal", "volume", 0.003785411784);
ge("qt", "volume", 0.003785411784 / 4);
ge("pt|us_pt", "volume", 0.003785411784 / 8);
ge("cup", "volume", 0.003785411784 / 16);
ge("oz", "volume", 0.003785411784 / 128);
ge("tbs", "volume", 0.003785411784 / 256);
ge("tsp", "volume", 0.003785411784 / 768);
ge("tspm", "volume", 5e-6);
ge("uk_gal", "volume", 0.00454609);
ge("uk_qt", "volume", 0.00454609 / 4);
ge("uk_pt", "volume", 0.00454609 / 8);
ge("barrel", "volume", 0.003785411784 * 42);
for (let [e, t] of [
  ["m", 1],
  ["in", 0.0254],
  ["ft", 0.3048],
  ["yd", 0.9144],
  ["mi", 1609.344],
  ["Nmi", 1852],
  ["ang", 1e-10],
])
  (ge(`${e}2|${e}^2`, "area", t ** 2, e === "m" || e === "ang", 0, 2),
    ge(`${e}3|${e}^3`, "volume", t ** 3, e === "m" || e === "ang", 0, 3));
ge("ar", "area", 100, true);
ge("ha", "area", 1e4);
ge("uk_acre", "area", 4046.8564224);
ge("us_acre", "area", (1200 / 3937) ** 2 * 43560);
ge("bit", "information", 1, true);
ge("byte", "information", 8, true);
ge("m/s|m/sec", "speed", 1, true);
ge("m/h|m/hr", "speed", 1 / 3600, true);
ge("mph", "speed", 1609.344 / 3600);
ge("kn", "speed", 1852 / 3600);
var _W = Object.freeze(nl.map((e) => Object.freeze(e))),
  eR = {
    Y: 24,
    Z: 21,
    E: 18,
    P: 15,
    T: 12,
    G: 9,
    M: 6,
    k: 3,
    h: 2,
    da: 1,
    e: 1,
    d: -1,
    c: -2,
    m: -3,
    u: -6,
    n: -9,
    p: -12,
    f: -15,
    a: -18,
    z: -21,
    y: -24,
  },
  A4 = { ki: 10, Mi: 20, Gi: 30, Ti: 40, Pi: 50, Ei: 60, Zi: 70, Yi: 80 };
function tR(e) {
  let t = nl.find((r) => r.symbol === e);
  if (t) return { ...t };
  for (let [r, o] of Object.entries(A4)) {
    if (!e.startsWith(r)) continue;
    let n = nl.find(
      (i) => i.symbol === e.slice(r.length) && i.dimension === "information",
    );
    if (n) return { ...n, symbol: e, factor: n.factor * 2 ** o };
  }
  for (let r of Object.keys(eR).sort((o, n) => n.length - o.length)) {
    if (!e.startsWith(r)) continue;
    let o = nl.find((n) => n.symbol === e.slice(r.length) && n.metric);
    if (o)
      return { ...o, symbol: e, factor: o.factor * 10 ** (eR[r] * o.power) };
  }
}
function rR(e, t, r) {
  if (!py.includes(e)) return hy;
  let o = (R = "#VALUE!") => {
    throw r.issue(R);
  };
  (t.length !== (e === "CONVERT" ? 3 : 2) || t.some((R) => R == null)) && o();
  let n = t.map(r.retrieve),
    i = 0,
    a = (R) => {
      if (Array.isArray(R)) {
        for (let C of R) a(C);
        return;
      }
      if ((++i > 4e4 && o("#LIMIT!"), r.isIssue(R))) throw R;
      typeof R == "number" && !Number.isFinite(R) && o("#NUM!");
    };
  n.forEach(a);
  let l = (R) => (Number.isFinite(R) ? R : o("#NUM!"));
  if (e === "CONVERT") {
    let [R, C, T] = n;
    (typeof C != "string" ||
      typeof T != "string" ||
      Array.isArray(R) ||
      R == null ||
      (typeof R == "string" && !R.trim())) &&
      o();
    let _ = Number(R);
    Number.isFinite(_) || o();
    let P = tR(C),
      S = tR(T);
    return (
      (!P || !S || P.dimension !== S.dimension) && o("#N/A"),
      P.offset === S.offset
        ? l(_ * (P.factor / S.factor))
        : l((_ * P.factor + P.offset - S.offset) / S.factor)
    );
  }
  let s = (R) => (Array.isArray(R) ? R.flat(1 / 0) : [R]),
    c = s(n[0]),
    u = s(n[1]);
  c.length !== u.length && o("#N/A");
  let f = c
      .map((R, C) => [R, u[C]])
      .filter((R) => R.every((C) => typeof C == "number")),
    d = f.length;
  (!d && e === "SLOPE" && o("#N/A"),
    d < (e === "STEYX" ? 3 : 2) && o("#DIV/0!"));
  let p = f.map((R) => R[0]),
    m = f.map((R) => R[1]),
    h = Zt(p),
    v = Zt(m),
    g = m.map((R) => l(R - v)),
    y = p.map((R) => l(R - h)),
    b = Math.max(...g.map(Math.abs)),
    x = Math.max(...y.map(Math.abs));
  if (((!b || (!x && (e === "RSQ" || e === "PEARSON"))) && o("#DIV/0!"), !x))
    return 0;
  let k = g.map((R) => R / b),
    w = y.map((R) => R / x),
    A = Nr(k.map((R) => R * R)),
    I = Nr(w.map((R) => R * R)),
    N = Nr(k.map((R, C) => R * w[C]));
  return e === "SLOPE"
    ? l((N / A) * (x / b))
    : e === "PEARSON"
      ? Math.max(-1, Math.min(1, N / Math.sqrt(A * I)))
      : e === "RSQ"
        ? Math.min(1, Math.max(0, (N / A) * (N / I)))
        : l(
            Math.sqrt(
              Nr(w.map((R, C) => (R - (N / A) * k[C]) ** 2)) / (d - 2),
            ) * x,
          );
}
var vy = [
    "COUNT",
    "COUNTA",
    "COUNTBLANK",
    "AVERAGE",
    "MIN",
    "MAX",
    "MEDIAN",
    "MODE",
    "MODE.SNGL",
    "STDEV.S",
    "STDEV.P",
    "VAR.S",
    "VAR.P",
    "PERCENTILE",
    "PERCENTILE.INC",
    "PERCENTILE.EXC",
    "QUARTILE",
    "QUARTILE.INC",
    "QUARTILE.EXC",
    "RANK",
    "RANK.EQ",
    "RANK.AVG",
    "AVERAGEA",
    "MAXA",
    "MINA",
    "STDEVA",
    "STDEVPA",
    "VARA",
    "VARPA",
    "SKEW",
    "SKEW.P",
    "TRIMMEAN",
    "MODE.MULT",
    "PERMUT",
    "PERMUTATIONA",
  ],
  yy = Symbol("not-statistical-function"),
  $f = (e) => (Array.isArray(e) ? e.flatMap($f) : [e]),
  N4 = (e) => !!e && typeof e == "object" && "ref" in e;
function gy(e) {
  let t = 0,
    r = 0;
  for (let o of e) {
    let n = t + o;
    ((r += Math.abs(t) >= Math.abs(o) ? t - n + o : o - n + t), (t = n));
  }
  return t + r;
}
function nR(e, t, r) {
  if (!vy.includes(e)) return yy;
  let o = (N = "#VALUE!") => {
      throw r.issue(N);
    },
    n = e.startsWith("RANK"),
    i = e.startsWith("PERCENTILE"),
    a = e.startsWith("QUARTILE"),
    l = e === "TRIMMEAN",
    s = e === "PERMUT" || e === "PERMUTATIONA",
    c = [
      "AVERAGEA",
      "MAXA",
      "MINA",
      "STDEVA",
      "STDEVPA",
      "VARA",
      "VARPA",
    ].includes(e);
  if (
    !t.length ||
    t.length > 255 ||
    (e === "COUNTBLANK" && t.length !== 1) ||
    ((i || a || l || s) && t.length !== 2) ||
    (n && (t.length < 2 || t.length > 3))
  )
    return o();
  let u = t.map(r.retrieve),
    f = 0,
    d = (N) => {
      if (Array.isArray(N)) {
        for (let R of N) d(R);
        return;
      }
      if (
        (++f > 1e5 && o("#LIMIT!"),
        r.isIssue(N) &&
          ["#LIMIT!", "#CYCLE!", "#BLOCKED!", "#ERROR!"].includes(
            String(N.error),
          ))
      )
        throw N;
    };
  for (let N of u) d(N);
  let p = (N) => (Number.isFinite(N) ? N : o("#NUM!")),
    m = (N, R) => {
      if (N === undefined && R !== undefined) return R;
      if (r.isIssue(N)) throw N;
      if (N == null) return 0;
      if (Array.isArray(N) || (typeof N == "string" && !N.trim())) return o();
      let C = Number(N);
      return Number.isFinite(C) ? C : o();
    };
  if (e === "COUNTA") return u.flatMap($f).filter((N) => N != null).length;
  if (e === "COUNTBLANK")
    return u.flatMap($f).filter((N) => N == null || N === "").length;
  let h = n ? 1 : 0;
  if (s) {
    let N = Math.trunc(m(u[0])),
      R = Math.trunc(m(u[1]));
    if (N < 0 || R < 0 || (e === "PERMUT" && R > N) || (!N && R > 0))
      return o("#NUM!");
    if (!R) return 1;
    if (e === "PERMUTATIONA") return p(N ** R);
    if (R > 170) return o("#NUM!");
    let C = 1;
    for (let T = 0; T < R; T++) C = p(C * (N - T));
    return C;
  }
  let v = n || i || a || l ? [u[h]] : u,
    g = [];
  for (let [N, R] of v.entries()) {
    let C = Array.isArray(R) || N4(t[n || i || a || l ? h : N]);
    for (let T of $f(R)) {
      if (r.isIssue(T)) {
        if (e === "COUNT") continue;
        throw T;
      }
      if (typeof T == "number") g.push(p(T));
      else if (c && C && T != null)
        if (typeof T == "boolean") g.push(Number(T));
        else if (typeof T == "string") g.push(0);
        else return o();
      else if (!C && T != null) {
        if (typeof T == "boolean") g.push(Number(T));
        else if (typeof T == "string" && T.trim() && Number.isFinite(Number(T)))
          g.push(Number(T));
        else if (e !== "COUNT") return o();
      }
    }
  }
  if (e === "COUNT") return g.length;
  if (e === "AVERAGE" || e === "AVERAGEA")
    return g.length ? p(Zt(g)) : o("#DIV/0!");
  if (e === "MIN" || e === "MINA")
    return g.length ? g.reduce((N, R) => Math.min(N, R)) : 0;
  if (e === "MAX" || e === "MAXA")
    return g.length ? g.reduce((N, R) => Math.max(N, R)) : 0;
  if (e.startsWith("STDEV") || e.startsWith("VAR")) {
    let N = e.endsWith(".S") || e === "STDEVA" || e === "VARA";
    if (g.length < (N ? 2 : 1)) return o("#DIV/0!");
    let R = Zt(g),
      C = g.map((S) => S - R),
      T = C.reduce((S, F) => Math.max(S, Math.abs(F)), 0);
    if (
      (Number.isFinite(T)
        ? T && (C = C.map((S) => S / T))
        : ((T = g.reduce((S, F) => Math.max(S, Math.abs(F)), 0)),
          (C = g.map((S) => S / T - R / T))),
      !T)
    )
      return 0;
    let _ = gy(C.map((S) => S * S)) / (g.length - (N ? 1 : 0)),
      P = Math.sqrt(_) * T;
    return p(e.startsWith("STDEV") ? P : P * P);
  }
  if (e === "SKEW" || e === "SKEW.P") {
    let N = g.length;
    if (N < 3) return o("#DIV/0!");
    let R = Zt(g),
      C = g.reduce((F, D) => Math.max(F, Math.abs(D)), 0);
    if (!C) return o("#DIV/0!");
    let T = g.map((F) => F / C - R / C),
      _ = gy(T.map((F) => F * F));
    if (!_) return o("#DIV/0!");
    let S = gy(T.map((F) => F ** 3)) / _ ** 1.5;
    return p(
      e === "SKEW.P" ? S * Math.sqrt(N) : (S * N * Math.sqrt(N - 1)) / (N - 2),
    );
  }
  if ((g.sort((N, R) => N - R), l)) {
    let N = m(u[1]);
    if (N < 0 || N >= 1 || !g.length) return o("#NUM!");
    let R = Math.floor((g.length * N) / 2);
    return p(Zt(g.slice(R, g.length - R)));
  }
  if (n) {
    let N = m(u[0]),
      R = m(u[2], 0),
      C = g.indexOf(N),
      T = g.lastIndexOf(N);
    return C < 0
      ? o("#N/A")
      : R
        ? 1 + (e === "RANK.AVG" ? (C + T) / 2 : C)
        : g.length - (e === "RANK.AVG" ? (C + T) / 2 : T);
  }
  if (e === "MODE.MULT") {
    let N = new Map(),
      R = 1;
    for (let T of g) {
      let _ = (N.get(T) ?? 0) + 1;
      (N.set(T, _), (R = Math.max(R, _)));
    }
    if (R === 1) return o("#N/A");
    let C = [...N].filter(([, T]) => T === R).map(([T]) => [T]);
    return C.length > 1e4 ? o("#LIMIT!") : C;
  }
  if (e === "MODE" || e === "MODE.SNGL") {
    let N = 0,
      R = 1,
      C = 0,
      T;
    for (let _ of g)
      ((C = _ === T ? C + 1 : 1), C > R && ((R = C), (N = _)), (T = _));
    return R > 1 ? N : o("#N/A");
  }
  if (!g.length) return o("#NUM!");
  if (e === "MEDIAN") {
    let N = Math.floor(g.length / 2);
    return g.length % 2 || g[N - 1] === g[N] ? g[N] : g[N - 1] / 2 + g[N] / 2;
  }
  let y = m(u[1]),
    b = a ? Math.trunc(y) / 4 : y,
    x = e.endsWith(".EXC");
  if (b < 0 || b > 1 || (x && (b === 0 || b === 1))) return o("#NUM!");
  let k = x ? b * (g.length + 1) - 1 : b * (g.length - 1);
  if (k < 0 || k > g.length - 1) return o("#NUM!");
  let w = Math.floor(k),
    A = Math.ceil(k),
    I = k - w;
  return g[w] === g[A] ? g[w] : p(g[w] * (1 - I) + g[A] * I);
}
var oR = ["LOOKUP", "CHOOSE"],
  by = Symbol("not an owned lookup function");
function iR(e, t, r) {
  if (!["VLOOKUP", "HLOOKUP", "LOOKUP", "CHOOSE"].includes(e)) return by;
  let o = (h = "#VALUE!") => {
      throw r.issue(h);
    },
    n = (h) => {
      if (r.isIssue(h)) throw h;
      return (Array.isArray(h) && o(), h);
    },
    i = (h) => {
      if (r.isIssue(h)) throw h;
      h === undefined && o();
      let v = Array.isArray(h) ? (h.every(Array.isArray) ? h : [h]) : [[h]];
      return (
        (!v.length ||
          !v[0].length ||
          v.some((g) => g.length !== v[0].length)) &&
          o(),
        v.length * v[0].length > 2e4 && o("#LIMIT!"),
        v
      );
    },
    a = (h) => (h.length !== 1 && h[0].length !== 1 && o(), h.flat());
  if (e === "CHOOSE") {
    (t.length < 2 || t.length > 255) && o();
    let h = n(r.retrieve(t[0]));
    typeof h == "string" && !h.trim() && o();
    let v = Math.trunc(Number(h));
    return ((!Number.isFinite(v) || v < 1 || v >= t.length) && o(), t[v] ?? 0);
  }
  (t.length < (e === "LOOKUP" ? 2 : 3) ||
    t.length > (e === "LOOKUP" ? 3 : 4)) &&
    o();
  let l = n(r.retrieve(t[0]));
  l === undefined && o();
  let s = i(r.retrieve(t[1])),
    c,
    u,
    f = true;
  if (e === "LOOKUP")
    t.length === 3
      ? ((c = a(s)), (u = a(i(r.retrieve(t[2])))), c.length !== u.length && o())
      : s[0].length > s.length
        ? ((c = s[0]), (u = s[s.length - 1]))
        : ((c = s.map((h) => h[0])), (u = s.map((h) => h[h.length - 1])));
  else {
    let h = n(r.retrieve(t[2]));
    typeof h == "string" && !h.trim() && o();
    let v = Math.trunc(Number(h));
    ((!Number.isFinite(v) || v < 1) && o(),
      v > (e === "VLOOKUP" ? s[0].length : s.length) && o("#REF!"));
    let g = t[3] == null ? true : n(r.retrieve(t[3]));
    (typeof g == "string" && !/^(true|false)$/i.test(g) && o(),
      (f = typeof g == "string" ? g.toLowerCase() === "true" : !!g),
      (c = e === "VLOOKUP" ? s.map((y) => y[0]) : s[0]),
      (u = e === "VLOOKUP" ? s.map((y) => y[v - 1]) : s[v - 1]));
  }
  let d = c.find(r.isIssue);
  if (d) throw d;
  let p = (h) => (typeof h == "string" ? h.toLowerCase() : h),
    m = -1;
  if (f) {
    (!["number", "string", "boolean"].includes(typeof l) ||
      c.some((g) => typeof g != typeof l)) &&
      o();
    let h = c.map(p),
      v = p(l);
    (h.some((g, y) => y > 0 && g < h[y - 1]) && o(),
      h.forEach((g, y) => {
        g <= v && (m = y);
      }));
  } else {
    let h = 0,
      v = typeof l == "string" && /[?*~]/.test(l);
    m = c.findIndex((g) =>
      v
        ? typeof g != "string"
          ? false
          : ((h += g.length * l.length),
            h > 1e5 && o("#LIMIT!"),
            r.matches(g, `=${l}`))
        : p(g) === p(l),
    );
  }
  return m < 0 ? o("#N/A") : (u[m] ?? 0);
}
var wy = Symbol("not an owned logical function");
function aR(e, t, r) {
  if (!["AND", "OR", "XOR", "NOT"].includes(e)) return wy;
  let o = () => {
    throw r.issue("#VALUE!");
  };
  (!t.length || t.length > 255 || (e === "NOT" && t.length !== 1)) && o();
  let n = [],
    i = 0,
    a = (l, s) => {
      if (++i > 1e5) throw r.issue("#LIMIT!");
      if (r.isIssue(l)) throw l;
      if (Array.isArray(l)) (e === "NOT" && o(), l.forEach((c) => a(c, true)));
      else if (typeof l == "boolean" || typeof l == "number") n.push(!!l);
      else if (l == null) s || n.push(false);
      else if (typeof l == "string") {
        if (s) return;
        (/^(true|false)$/i.test(l) || o(), n.push(l.toLowerCase() === "true"));
      } else o();
    };
  return (
    t.forEach((l) =>
      a(
        r.retrieve(l),
        e !== "NOT" && !!l && typeof l == "object" && "ref" in l,
      ),
    ),
    n.length || o(),
    e === "NOT"
      ? !n[0]
      : e === "AND"
        ? n.every(Boolean)
        : e === "OR"
          ? n.some(Boolean)
          : n.filter(Boolean).length % 2 === 1
  );
}
var Ey = [
    "XLOOKUP",
    "XMATCH",
    "IFS",
    "SWITCH",
    "TEXTJOIN",
    "CONCAT",
    "MAXIFS",
    "MINIFS",
    "FILTER",
    "SORT",
    "SORTBY",
    "UNIQUE",
    "SEQUENCE",
    "TRANSPOSE",
    "TAKE",
    "DROP",
    "CHOOSECOLS",
    "CHOOSEROWS",
    "HSTACK",
    "VSTACK",
    "TOCOL",
    "TOROW",
    "EXPAND",
    "TEXTBEFORE",
    "TEXTAFTER",
    "TEXTSPLIT",
    "STDEV.P",
    "VAR.P",
    "PERCENTILE.INC",
    "PERCENTILE.EXC",
    "QUARTILE.INC",
    "QUARTILE.EXC",
    "RANK.EQ",
    "RANK.AVG",
    "MODE.SNGL",
    "MDETERM",
    "MINVERSE",
  ],
  ol = [
    "CHOOSE",
    "FILTER",
    "SORT",
    "SORTBY",
    "UNIQUE",
    "SEQUENCE",
    "TRANSPOSE",
    "TAKE",
    "DROP",
    "CHOOSECOLS",
    "CHOOSEROWS",
    "HSTACK",
    "VSTACK",
    "TOCOL",
    "TOROW",
    "EXPAND",
    "TEXTSPLIT",
    "MUNIT",
    "MMULT",
    "MINVERSE",
    "FREQUENCY",
    "GROWTH",
    "LINEST",
    "LOGEST",
    "TREND",
    "MODE.MULT",
    "XLOOKUP",
  ],
  Mf = Symbol("not a modern function"),
  Lr = (e) => (Array.isArray(e) ? e.flatMap(Lr) : [e]),
  _t = (e) => (Array.isArray(e) ? (e.every(Array.isArray) ? e : [e]) : [[e]]),
  Hn = (e) =>
    e == null ? "" : typeof e == "boolean" ? (e ? "TRUE" : "FALSE") : String(e),
  Of = (e, t) =>
    typeof e == "number" && typeof t == "number"
      ? e - t
      : Hn(e).toLowerCase().localeCompare(Hn(t).toLowerCase()),
  sR = (e, t) =>
    typeof e == "string" && typeof t == "string"
      ? e.toLowerCase() === t.toLowerCase()
      : e === t,
  lR = (e, t, r) => {
    if (e == null || t == null) return e == null ? (t == null ? 0 : 1) : -1;
    let o = (n) => (typeof n == "number" ? 0 : typeof n == "string" ? 1 : 2);
    return (o(e) - o(t) || Of(e, t)) * r;
  };
function cR(e, t, r) {
  if (!Ey.includes(e)) return Mf;
  let o = (f = "#VALUE!") => {
      throw r.issue(f);
    },
    n = (f, d) => {
      if (r.isIssue(f)) throw f;
      if (f === undefined && d !== undefined) return d;
      if (Array.isArray(f) || (typeof f == "string" && !f.trim())) return o();
      let p = Number(f);
      return Number.isFinite(p) ? p : o();
    },
    i = (f, d) => Math.trunc(n(f, d)),
    a = (f, d) => {
      ((f < 1 || d < 1) && o("#CALC!"), f * d > 1e4 && o("#LIMIT!"));
    },
    l = (f) => {
      if (r.isIssue(f)) throw f;
      return (
        Array.isArray(f) && o(),
        typeof f == "string" && !/^(true|false)$/i.test(f) && o(),
        typeof f == "string" ? f.toLowerCase() === "true" : !!f
      );
    };
  if (e === "IFS") {
    (t.length < 2 || t.length % 2) && o();
    for (let f = 0; f < t.length; f += 2) {
      let d = r.retrieve(t[f]);
      if ((Array.isArray(d) && o(), l(d))) return r.retrieve(t[f + 1]) ?? 0;
    }
    return o("#N/A");
  }
  if (e === "SWITCH") {
    t.length < 3 && o();
    let f = r.retrieve(t[0]);
    if (r.isIssue(f)) return f;
    Array.isArray(f) && o();
    for (let d = 1; d + 1 < t.length; d += 2) {
      let p = r.retrieve(t[d]);
      if (r.isIssue(p)) return p;
      if ((Array.isArray(p) && o(), sR(f, p))) return r.retrieve(t[d + 1]) ?? 0;
    }
    return t.length % 2 === 0 ? (r.retrieve(t.at(-1)) ?? 0) : o("#N/A");
  }
  let s = t.map(r.retrieve),
    u = {
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
    }[e];
  if (
    (u && (t.length < u[0] || t.length > u[1]) && o(),
    u && e !== "SEQUENCE" && s[0] === undefined && o(),
    u && !["HSTACK", "VSTACK"].includes(e) && r.isIssue(s[0]))
  )
    throw s[0];
  if (!ol.includes(e) && e !== "MAXIFS" && e !== "MINIFS") {
    let f = s.flatMap(Lr).find(r.isIssue);
    if (f) return f;
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
    ].includes(e)
  ) {
    let f = e.startsWith("RANK"),
      d = (
        e === "STDEV.P" || e === "VAR.P" || e === "MODE.SNGL"
          ? s.flatMap(Lr)
          : Lr(s[f ? 1 : 0])
      )
        .filter((g) => typeof g == "number")
        .sort((g, y) => g - y);
    if ((d.length || o("#NUM!"), f)) {
      let g = n(s[0]),
        y = n(s[2], 0),
        b = d.indexOf(g),
        x = d.lastIndexOf(g);
      return (
        b < 0 && o("#N/A"),
        y
          ? 1 + (e === "RANK.AVG" ? (b + x) / 2 : b)
          : d.length - (e === "RANK.AVG" ? (b + x) / 2 : x)
      );
    }
    if (e === "MODE.SNGL") {
      let g = new Map();
      d.forEach((b) => g.set(b, (g.get(b) ?? 0) + 1));
      let y = [...g].sort((b, x) => x[1] - b[1] || b[0] - x[0]);
      return y[0][1] > 1 ? y[0][0] : o("#N/A");
    }
    if (e === "STDEV.P" || e === "VAR.P") {
      let g = d.reduce((b, x) => b + x, 0) / d.length,
        y = d.reduce((b, x) => b + (x - g) ** 2, 0) / d.length;
      return e === "STDEV.P" ? Math.sqrt(y) : y;
    }
    let p = n(s[1]) / (e.startsWith("QUARTILE") ? 4 : 1),
      m = e.endsWith("EXC"),
      h = m ? p * (d.length + 1) - 1 : p * (d.length - 1);
    (p < 0 || p > 1 || h < 0 || h > d.length - 1) && o("#NUM!");
    let v = Math.floor(h);
    return d[v] + (h - v) * ((d[v + 1] ?? d[v]) - d[v]);
  }
  if (e === "MDETERM" || e === "MINVERSE") {
    t.length !== 1 && o();
    let f = _t(s[0]),
      d = f.length;
    (d > 32 && o("#LIMIT!"), (!d || f.some((v) => v.length !== d)) && o());
    for (let v of f.flat()) {
      if (r.isIssue(v)) return v;
      (typeof v != "number" && o(), Number.isFinite(v) || o("#NUM!"));
    }
    let p = f.map((v, g) => [
        ...v.map((y) => y),
        ...Array.from({ length: d }, (y, b) => (g === b ? 1 : 0)),
      ]),
      m = 1;
    for (let v = 0; v < d; v++) {
      let g = v;
      for (let b = v + 1; b < d; b++)
        Math.abs(p[b][v]) > Math.abs(p[g][v]) && (g = b);
      if (p[g][v] === 0) return e === "MDETERM" ? 0 : o("#NUM!");
      g !== v && (([p[v], p[g]] = [p[g], p[v]]), (m *= -1));
      let y = p[v][v];
      ((m *= y), (p[v] = p[v].map((b) => b / y)));
      for (let b = 0; b < d; b++)
        if (b !== v) {
          let x = p[b][v];
          p[b] = p[b].map((k, w) => k - x * p[v][w]);
        }
    }
    let h = p.map((v) => v.slice(d));
    return e === "MDETERM"
      ? Number.isFinite(m)
        ? m
        : o("#NUM!")
      : (h.some((v) => v.some((g) => !Number.isFinite(g))) && o("#NUM!"), h);
  }
  if (e === "CONCAT" || e === "TEXTJOIN") {
    e === "TEXTJOIN" && s.length < 3 && o();
    let f = e === "TEXTJOIN" ? Hn(s[0]) : "",
      d = e === "TEXTJOIN" && l(s[1]),
      p = s
        .slice(e === "TEXTJOIN" ? 2 : 0)
        .flatMap(Lr)
        .filter((m) => !d || (m !== null && m !== ""));
    return (
      p.reduce((m, h) => m + Hn(h).length, 0) +
        Math.max(0, p.length - 1) * f.length >
        32767 && o("#LIMIT!"),
      p.map(Hn).join(f)
    );
  }
  if (e === "MAXIFS" || e === "MINIFS") {
    (s.length < 3 || s.length % 2 === 0) && o();
    let f = Lr(s[0]),
      d = [];
    for (let v = 1; v < s.length; v += 2) {
      let g = Lr(s[v]);
      if (
        ((g.length !== f.length ||
          _t(s[v]).length !== _t(s[0]).length ||
          _t(s[v])[0].length !== _t(s[0])[0].length ||
          Array.isArray(s[v + 1])) &&
          o(),
        r.isIssue(s[v + 1]))
      )
        throw s[v + 1];
      d.push([g, s[v + 1]]);
    }
    f.length * d.length > 1e5 && o("#LIMIT!");
    let p = f.filter((v, g) =>
        d.every(([y, b]) => !r.isIssue(y[g]) && r.matches(y[g], b)),
      ),
      m = p.find(r.isIssue);
    if (m) return m;
    let h = p.filter((v) => typeof v == "number");
    return h.length
      ? h.reduce((v, g) => (e === "MAXIFS" ? Math.max(v, g) : Math.min(v, g)))
      : 0;
  }
  if (e === "XMATCH" || e === "XLOOKUP") {
    ((t.length < (e === "XMATCH" ? 2 : 3) ||
      t.length > (e === "XMATCH" ? 4 : 6)) &&
      o(),
      (s[0] === undefined ||
        s[1] === undefined ||
        (e === "XLOOKUP" && s[2] === undefined)) &&
        o());
    let f = s[0],
      d = _t(s[1]);
    if (r.isIssue(f)) return f;
    if (r.isIssue(s[1])) return s[1];
    (Array.isArray(f) ||
      !d.length ||
      !d[0].length ||
      d.some((k) => k.length !== d[0].length) ||
      (d.length > 1 && d[0].length > 1)) &&
      o();
    let p = Lr(d),
      m = n(s[e === "XMATCH" ? 2 : 4], 0),
      h = n(s[e === "XMATCH" ? 3 : 5], 1);
    (![0, -1, 1, 2].includes(m) || ![1, -1, 2, -2].includes(h)) && o();
    let v = p.find(r.isIssue);
    if (v) return v;
    let g = [];
    if (e === "XLOOKUP") {
      if (r.isIssue(s[2])) return s[2];
      ((g = _t(s[2])),
        (!g.length ||
          !g[0].length ||
          g.some((k) => k.length !== g[0].length)) &&
          o(),
        (d.length > 1 ? g.length !== p.length : g[0].length !== p.length) &&
          o());
    }
    (Math.abs(h) === 2 || m === -1 || m === 1) &&
      (!["number", "string", "boolean"].includes(typeof f) ||
        p.some((k) => typeof k != typeof f)) &&
      o();
    let y = p.map((k, w) => w);
    (h < 0 && y.reverse(),
      Math.abs(h) === 2 &&
        p.some((k, w) => w > 0 && Of(k, p[w - 1]) * Math.sign(h) < 0) &&
        o());
    let b = 0,
      x = y.find((k) =>
        m !== 2 || typeof f != "string"
          ? sR(p[k], f)
          : typeof p[k] != "string"
            ? false
            : ((b += p[k].length * f.length),
              b > 1e5 && o("#LIMIT!"),
              r.matches(p[k], `=${f}`)),
      );
    if (x === undefined && (m === -1 || m === 1)) {
      let k = y.filter((w) => Math.sign(Of(p[w], f)) === m);
      (k.sort((w, A) => Of(p[w], p[A]) * m), (x = k[0]));
    }
    return x === undefined
      ? e === "XLOOKUP" && s[3] !== undefined
        ? (s[3] ?? 0)
        : o("#N/A")
      : e === "XMATCH"
        ? x + 1
        : d.length > 1
          ? g[x].length === 1
            ? (g[x][0] ?? 0)
            : [g[x]]
          : g.length === 1
            ? (g[0][x] ?? 0)
            : g.map((k) => [k[x]]);
  }
  if (e === "SEQUENCE") {
    s[0] === undefined && s[1] === undefined && o();
    let f = i(s[0], 1),
      d = i(s[1], 1),
      p = n(s[2], 1),
      m = n(s[3], 1);
    return (
      a(f, d),
      Array.from({ length: f }, (h, v) =>
        Array.from({ length: d }, (g, y) => p + (v * d + y) * m),
      )
    );
  }
  if (e === "FILTER") {
    let f = _t(s[0]),
      d = _t(s[1]);
    s[1] === undefined && o();
    let p = Lr(d).find(r.isIssue);
    if (p) throw p;
    let m;
    if (d.length === f.length && d[0].length === 1)
      m = f.filter((h, v) => l(d[v][0]));
    else if (d.length === 1 && d[0].length === f[0].length)
      m = f.map((h) => h.filter((v, g) => l(d[0][g])));
    else return o();
    return !m.length || !m[0].length
      ? s[2] !== undefined
        ? _t(s[2])
        : o("#CALC!")
      : m;
  }
  if (e === "TRANSPOSE")
    return _t(s[0])[0].map((f, d) => _t(s[0]).map((p) => p[d]));
  if (e === "SORT" || e === "SORTBY") {
    let f = _t(s[0]).map((g) => [...g]);
    if (e === "SORTBY") {
      let g = _t(s[1]);
      s[1] === undefined && o();
      let y =
          g.length === 1 &&
          g[0].length === f[0].length &&
          (f.length === 1 || g[0].length > 1),
        b = y ? f[0].length : f.length,
        x = [];
      for (let w = 1; w < s.length; w += 2) {
        s[w] === undefined && o();
        let A = _t(s[w]),
          I = n(s[w + 1], 1);
        ((y
          ? A.length !== 1 || A[0].length !== b
          : A.length !== b || A.some((C) => C.length !== 1)) ||
          ![-1, 1].includes(I)) &&
          o();
        let N = Lr(A),
          R = N.find(r.isIssue);
        if (R) throw R;
        x.push({ data: N, order: I });
      }
      let k = Array.from({ length: b }, (w, A) => A).sort((w, A) => {
        for (let I of x) {
          let N = lR(I.data[w], I.data[A], I.order);
          if (N) return N;
        }
        return w - A;
      });
      return y ? f.map((w) => k.map((A) => w[A])) : k.map((w) => f[w]);
    }
    let d = i(s[1], 1) - 1,
      p = n(s[2], 1),
      m = s[3] === undefined ? false : l(s[3]);
    [-1, 1].includes(p) || o();
    let h = m ? f[0].map((g, y) => f.map((b) => b[y])) : f;
    (d < 0 || d >= h[0].length) && o();
    let v = h.map((g) => g[d]).find(r.isIssue);
    if (v) throw v;
    return (
      h.sort((g, y) => lR(g[d], y[d], p)),
      m ? h[0].map((g, y) => h.map((b) => b[y])) : h
    );
  }
  if (e === "UNIQUE") {
    let f = _t(s[0]),
      d = s[1] === undefined ? false : l(s[1]),
      p = s[2] === undefined ? false : l(s[2]),
      m = d ? f[0].map((b, x) => f.map((k) => k[x])) : f,
      h = (b) =>
        JSON.stringify(
          b.map((x) => (typeof x == "string" ? x.toLowerCase() : (x ?? 0))),
        ),
      v = new Map();
    m.forEach((b) => v.set(h(b), (v.get(h(b)) ?? 0) + 1));
    let g = new Set(),
      y = m.filter((b) => {
        let x = h(b);
        return g.has(x) || (p && v.get(x) !== 1) ? false : (g.add(x), true);
      });
    return (
      y.length || o("#CALC!"),
      d ? y[0].map((b, x) => y.map((k) => k[x])) : y
    );
  }
  if (e === "TAKE" || e === "DROP") {
    let f = _t(s[0]),
      d = s[1] === undefined ? undefined : i(s[1]),
      p = s[2] === undefined ? undefined : i(s[2]),
      m = (v, g) =>
        g === undefined
          ? v
          : e === "TAKE"
            ? g > 0
              ? v.slice(0, g)
              : v.slice(g)
            : g >= 0
              ? v.slice(g)
              : v.slice(0, g);
    (d === undefined && p === undefined && o(),
      (d === 0 || p === 0) && o("#CALC!"));
    let h = m(f, d).map((v) => m(v, p));
    return ((!h.length || !h[0].length) && o("#CALC!"), h);
  }
  if (e === "CHOOSECOLS" || e === "CHOOSEROWS") {
    let f = _t(s[0]),
      d = e === "CHOOSECOLS" ? f[0].length : f.length,
      p = s
        .slice(1)
        .flatMap(Lr)
        .map((m) => {
          let h = i(m),
            v = h > 0 ? h - 1 : d + h;
          return ((!h || v < 0 || v >= d) && o(), v);
        });
    return (
      p.length || o(),
      e === "CHOOSECOLS" ? f.map((m) => p.map((h) => m[h])) : p.map((m) => f[m])
    );
  }
  if (e === "HSTACK" || e === "VSTACK") {
    let f = s.map(_t),
      d =
        e === "VSTACK"
          ? f.reduce((m, h) => m + h.length, 0)
          : Math.max(...f.map((m) => m.length)),
      p =
        e === "HSTACK"
          ? f.reduce((m, h) => m + h[0].length, 0)
          : Math.max(...f.map((m) => m[0].length));
    return (
      a(d, p),
      e === "VSTACK"
        ? f.flatMap((m) =>
            m.map((h) =>
              Array.from({ length: p }, (v, g) =>
                g < h.length ? h[g] : r.issue("#N/A"),
              ),
            ),
          )
        : Array.from({ length: d }, (m, h) =>
            f.flatMap((v) =>
              Array.from({ length: v[0].length }, (g, y) =>
                h < v.length ? v[h][y] : r.issue("#N/A"),
              ),
            ),
          )
    );
  }
  if (e === "TOCOL" || e === "TOROW") {
    let f = _t(s[0]),
      d = i(s[1], 0),
      p = s[2] === undefined ? false : l(s[2]);
    [0, 1, 2, 3].includes(d) || o();
    let m = (p ? f[0].flatMap((h, v) => f.map((g) => g[v])) : f.flat()).filter(
      (h) =>
        !((d === 1 || d === 3) && h === null) &&
        !((d === 2 || d === 3) && r.isIssue(h)),
    );
    return (
      m.length || o("#CALC!"),
      e === "TOCOL" ? m.map((h) => [h ?? 0]) : [m.map((h) => h ?? 0)]
    );
  }
  if (e === "EXPAND") {
    let f = _t(s[0]),
      d = i(s[1], f.length),
      p = i(s[2], f[0].length);
    return (
      a(d, p),
      (d < f.length || p < f[0].length) && o(),
      Array.from({ length: d }, (m, h) =>
        Array.from({ length: p }, (v, g) =>
          h < f.length && g < f[h].length ? f[h][g] : (s[3] ?? r.issue("#N/A")),
        ),
      )
    );
  }
  if (e === "TEXTBEFORE" || e === "TEXTAFTER") {
    let f = Hn(s[0]),
      d = Hn(s[1]),
      p = i(s[2], 1),
      m = i(s[3], 0) === 1,
      h = i(s[4], 0) === 1;
    (!p || !d) && o();
    let v = m ? f.toLowerCase() : f,
      g = m ? d.toLowerCase() : d,
      y = [];
    for (let x = 0; x <= v.length - g.length;) {
      let k = v.indexOf(g, x);
      if (k < 0) break;
      (y.push(k), (x = k + g.length));
    }
    let b = y[p > 0 ? p - 1 : y.length + p];
    return b === undefined
      ? h && Math.abs(p) === y.length + 1
        ? e === "TEXTBEFORE"
          ? p > 0
            ? f
            : ""
          : p > 0
            ? ""
            : f
        : s[5] !== undefined
          ? s[5]
          : o("#N/A")
      : e === "TEXTBEFORE"
        ? f.slice(0, b)
        : f.slice(b + d.length);
  }
  if (e === "TEXTSPLIT") {
    (s.length < 2 || s.length > 6) && o();
    let f = s.slice(0, 5).flatMap(Lr).find(r.isIssue);
    if (f) return f;
    let d = Hn(s[0]),
      p = s[3] === undefined ? false : l(s[3]),
      m = n(s[4], 0);
    (Array.isArray(s[0]) || ![0, 1].includes(m)) && o();
    let h = (k) =>
        Lr(k)
          .filter((w) => w != null)
          .map((w) => {
            if (r.isIssue(w)) throw w;
            return Hn(w);
          })
          .filter(Boolean),
      v = h(s[1]),
      g = h(s[2]);
    (!v.length && !g.length && o(),
      (v.length + g.length > 100 || [...v, ...g].join("").length > 1e4) &&
        o("#LIMIT!"));
    let y = (k, w) =>
        w.length
          ? k
              .split(
                new RegExp(
                  w
                    .map((A) => A.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"))
                    .join("|"),
                  m ? "i" : "",
                ),
              )
              .filter((A) => !p || A !== "")
          : [k],
      b = y(d, g).map((k) => y(k, v)),
      x = Math.max(...b.map((k) => k.length));
    return (
      a(b.length, x),
      b.map((k) =>
        Array.from({ length: x }, (w, A) => k[A] ?? s[5] ?? r.issue("#N/A")),
      )
    );
  }
  return Mf;
}
var R4 = (e) => `${e.scopeSheetId ?? ""}:${e.name.toLowerCase()}`,
  T4 = (e) => `'${e.replace(/'/g, "''")}'!`;
function C4(e, t = false) {
  let r = [],
    o = 0,
    n = (a) => {
      if (t) throw Error(a);
    },
    i = () => {
      let a = 0;
      do (e[o] === "[" && a++, e[o] === "]" && a--, o++);
      while (o < e.length && a);
      a && n("Unbalanced table references cannot be safely migrated.");
    };
  for (; o < e.length;) {
    let a = o,
      l = e[o];
    if (l === '"') {
      o++;
      let m = false;
      for (; o < e.length;)
        if (e[o++] === '"')
          if (e[o] === '"') o++;
          else {
            m = true;
            break;
          }
      m || n("An unterminated formula string cannot be safely migrated.");
      continue;
    }
    if (l === "#") {
      let m =
        /^#(?:REF!|NAME\?|VALUE!|DIV\/0!|N\/A|NUM!|NULL!|SPILL!|CALC!|CYCLE!|LIMIT!|BLOCKED!|ERROR!)/i.exec(
          e.slice(o),
        );
      if (m) {
        o += m[0].length;
        continue;
      }
      o++;
      continue;
    }
    let s =
      /^(?:'(?:[^']|'')+'|[A-Za-z0-9_.]+):(?:'(?:[^']|'')+'|[A-Za-z0-9_.]+)!/.exec(
        e.slice(o),
      );
    if (s) {
      (n("Three-dimensional references cannot be safely migrated."),
        (o += s[0].length));
      continue;
    }
    let c;
    if (l === "'") {
      let m = /^'((?:[^']|'')+)'!/.exec(e.slice(o));
      if (!m) {
        for (
          n("An unsupported quoted reference cannot be safely migrated."), o++;
          o < e.length && e[o++] !== "'";
        );
        continue;
      }
      ((c = m[1].replace(/''/g, "'")),
        (o += m[0].length),
        /[\[\]:|]/.test(c) &&
          n(
            "External or three-dimensional references cannot be safely migrated.",
          ));
    } else {
      let m = /^([A-Za-z0-9_\\][A-Za-z0-9_.\\]*)!/.exec(e.slice(o));
      m && ((c = m[1]), (o += m[0].length));
    }
    if (c === undefined && l === "[") {
      (e.startsWith("[@", o) ||
        n(
          "External or unqualified bracket references cannot be safely migrated.",
        ),
        i());
      continue;
    }
    let u =
      /^(?:\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}|\$?[1-9]\d*:\$?[1-9]\d*)(?![A-Za-z0-9_.\\])/.exec(
        e.slice(o),
      );
    if (u) {
      ((o += u[0].length),
        r.push({
          start: a,
          end: o,
          text: e.slice(a, o),
          name: u[0],
          ...(c ? { qualifier: c } : {}),
          cell: true,
        }));
      continue;
    }
    if (c === undefined && /[0-9.]/.test(l)) {
      let m = /^(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?/.exec(e.slice(o));
      if (m) {
        o += m[0].length;
        continue;
      }
    }
    let f =
      /^(?:\$?[A-Za-z]{1,3}\$?[1-9]\d*(?![A-Za-z0-9_.\\])|[A-Za-z_\\][A-Za-z0-9_.\\]*)/.exec(
        e.slice(o),
      );
    if (!f) {
      (c && n("The qualified reference cannot be safely migrated."),
        l === "!" && n("The reference qualifier cannot be safely migrated."),
        l === "@" && e[o + 1] === "[" ? (o++, i()) : o++);
      continue;
    }
    let d = f[0];
    if (((o += d.length), e[o] === "[")) {
      i();
      continue;
    }
    let p = /^\s*\(/.test(e.slice(o));
    (p &&
      /^(?:_xlfn\.)?(?:INDIRECT|EVALUATE|GET\.CELL)$/i.test(d) &&
      n(
        `${d} uses dynamic or locally bound names and cannot be safely migrated.`,
      ),
      !(a > 0 && /[A-Za-z0-9_.$\\]/.test(e[a - 1])) &&
        r.push({
          start: a,
          end: o,
          text: e.slice(a, o),
          name: d,
          ...(p ? { call: true } : {}),
          ...(c ? { qualifier: c } : {}),
          ...(/^\$?[A-Za-z]{1,3}\$?[1-9]\d*$/.test(d) ? { cell: true } : {}),
        }));
  }
  return r;
}
function uR(e, t = false) {
  let r = C4(e, t);
  if (!/\b(?:_xlfn\.)?(?:LET|LAMBDA)\s*\(/i.test(e)) return r;
  let o = new Map();
  try {
    Pn(Gi(e), (n, i, a, l) => {
      o.set(n.start, {
        bound: new Set(i),
        excluded: l || (!n.qualifier && i.has(vn(n.name))),
      });
    });
  } catch (n) {
    if (t) throw n;
    return r;
  }
  return r
    .filter((n) => !o.get(n.start)?.excluded)
    .map((n) => ({ ...n, bound: o.get(n.start)?.bound }));
}
function S4(e, t, r) {
  if (e.cell || (e.call && !e.qualifier && Pf.includes(ks(e.name)))) return;
  let o = e.qualifier
    ? t.sheets.find((n) => n.name.toLowerCase() === e.qualifier.toLowerCase())
        ?.id
    : r;
  if (o)
    return (
      t.definedNames?.find(
        (n) =>
          n.scopeSheetId === o && n.name.toLowerCase() === e.name.toLowerCase(),
      ) ??
      t.definedNames?.find(
        (n) => !n.scopeSheetId && n.name.toLowerCase() === e.name.toLowerCase(),
      )
    );
}
function fR(e, t, r) {
  let o = "",
    n = 0;
  for (let i of t) ((o += e.slice(n, i.start) + r(i)), (n = i.end));
  return o + e.slice(n);
}
function il(e, t, r, o = [], n) {
  if (o.length > 32) throw new Error("#LIMIT!");
  let i = fR(e, uR(e), (a) => {
    if (a.call) return a.text;
    let l = S4(a, t, r);
    if (!l || ex(l.refersTo)) return a.text;
    let s = R4(l);
    if (o.includes(s)) throw new Error("#CYCLE!");
    let c =
      l.scopeSheetId ??
      (a.qualifier
        ? t.sheets.find(
            (f) => f.name.toLowerCase() === a.qualifier.toLowerCase(),
          ).id
        : r);
    if (n && !n(c)) throw new Error("#BLOCKED!");
    let u = il(l.refersTo.slice(1), t, c, [...o, s], n);
    if (c !== r) {
      let f = t.sheets.find((p) => p.id === c),
        d = uR(u);
      u = fR(u, d, (p) => {
        let m = d.indexOf(p),
          h = d[m - 1],
          v = h?.cell && u.slice(h.end, p.start).trim() === ":";
        return p.cell && !p.qualifier && !v ? T4(f.name) + p.text : p.text;
      });
    }
    return `(${u})`;
  });
  if (i.length > 16384) throw new Error("#LIMIT!");
  return i;
}
function xy(e) {
  let t = "",
    r = 0;
  for (; r < e.length;) {
    let o = e[r];
    if (o === '"' || o === "'") {
      let i = r++,
        a = o;
      for (; r < e.length;)
        if (e[r++] === a)
          if (e[r] === a) r++;
          else break;
      t += e.slice(i, r);
      continue;
    }
    let n = /^(?:[A-Za-z_][A-Za-z0-9_.]*\[|\[@)/.exec(e.slice(r));
    if (n && (r === 0 || !/[A-Za-z0-9_.$]/.test(e[r - 1]))) {
      let i = r;
      r += n[0].length;
      let a = 1;
      for (; r < e.length && a;)
        (e[r] === "[" && a++, e[r] === "]" && a--, r++);
      let l = e.slice(i, r);
      !a && !/[!|]/.test(l) && !/^[A-Za-z0-9_!]/.test(e.slice(r))
        ? (t += "0")
        : (t += l);
      continue;
    }
    t += e[r++];
  }
  return t;
}
var Nn = (e) => {
  throw Error(
    `XLSX defined names: ${e} Import/export stopped to prevent changed formula bindings.`,
  );
};
function Iy(e) {
  if (!e.startsWith("=") || e.length < 2 || e.length > 4096)
    return Nn(
      "A definition must contain a formula of 1\u20134,095 characters after =.",
    );
  let t = 0,
    r = 0,
    o = 0,
    n = 1,
    i = "";
  for (; n < e.length;) {
    let l = e[n];
    if (!r && (l === '"' || l === "'")) {
      let s = l,
        c = n++,
        u = false;
      for (; n < e.length;)
        if (e[n++] === s)
          if (e[n] === s) n++;
          else {
            u = true;
            break;
          }
      if (!u) return Nn("Unterminated formula string or worksheet qualifier.");
      if (s === "'") {
        if (e[n] !== "!") return Nn("Malformed quoted worksheet qualifier.");
        i += e.slice(c, n);
      } else i += '""';
      continue;
    }
    if ((l === "[" && r++, l === "]" && --r < 0))
      return Nn("Unbalanced bracket reference.");
    if (!r) {
      if ((l === "{" && o++, l === "}" && --o < 0))
        return Nn("Unbalanced array braces.");
      if ((l === "(" && t++, l === ")" && --t < 0))
        return Nn("Unbalanced formula parentheses.");
    }
    ((i += l), n++);
  }
  if (t || r || o)
    return Nn(
      "Unbalanced formula parentheses, array braces or bracket reference.",
    );
  let a = xy(i);
  if (
    /[\[\]|]/.test(a) ||
    /(?:https?:|file:|javascript:|data:|\\\\)/i.test(a) ||
    /\b(?:_xlfn\.)?(?:WEBSERVICE|HYPERLINK|IMAGE|RTD|DDE|CALL|EXEC|EXECUTE|REGISTER(?:\.ID)?|RUN|SQL\.REQUEST|STOCKHISTORY|EVALUATE|GET\.CELL)\s*\(/i.test(
      a,
    )
  )
    return Nn("External or executable name formulas are not supported.");
  if (
    /(?:'(?:[^']|'')+'|[A-Za-z0-9_.]+):(?:'(?:[^']|'')+'|[A-Za-z0-9_.]+)!/.test(
      a,
    )
  )
    return Nn("Three-dimensional name references are not supported.");
  if (!i.trim()) return Nn("An empty formula is not supported.");
}
var dR = ["LET", "LAMBDA"],
  ba = { depth: 64, calls: 1024, work: 1e6, retainedValues: 1e5 },
  ky = (e) =>
    !!e && typeof e == "object" && "kind" in e && e.kind === "formula-closure",
  Uf = "__LGVALUE";
function mR(e, t, r, o) {
  let n = 0,
    i = 0,
    a = 0,
    l = new Map(),
    s = new Set(),
    c = (x, k) => {
      throw (o.diagnostic(x, k), o.issue(x));
    },
    u = (x, k) => {
      ((n += x),
        (k > ba.depth || n > ba.work) &&
          c(
            "#LIMIT!",
            `Reusable formula evaluation exceeds ${ba.depth} levels or ${ba.work.toLocaleString()} work units.`,
          ));
    },
    f = (x) => {
      let k = Array.isArray(x)
        ? x.reduce((w, A) => w + (Array.isArray(A) ? A.length : 1), 0)
        : 1;
      return (
        (a += k),
        a > ba.retainedValues &&
          c(
            "#LIMIT!",
            "Reusable formula intermediate values exceed 100,000 cells.",
          ),
        x
      );
    },
    d = (x) => {
      let k = l.get(x);
      if (!k) {
        u(x.length, 0);
        try {
          try {
            Iy(x.startsWith("=") ? x : `=${x}`);
          } catch (w) {
            let A = w instanceof Error ? w.message : "Unsafe formula.";
            c(
              /External|executable|Three-dimensional/i.test(A)
                ? "#BLOCKED!"
                : "#VALUE!",
              A,
            );
          }
          ((k = Gi(x)),
            Pn(k, (w, A, I, N) => {
              (!N &&
                /^__LG/i.test(w.name) &&
                c(
                  "#NAME?",
                  "Internal calculation names cannot be called from a workbook.",
                ),
                /^(?:_xlfn\.)?(?:WEBSERVICE|HYPERLINK|IMAGE|RTD|DDE|CALL|EXEC|EXECUTE|REGISTER(?:\.ID)?|RUN|SQL\.REQUEST|STOCKHISTORY|EVALUATE|GET\.CELL)$/i.test(
                  w.name,
                ) &&
                  c(
                    "#BLOCKED!",
                    "External or executable custom functions are not supported.",
                  ));
            }));
        } catch (w) {
          if (o.isIssue(w)) throw w;
          c(
            w instanceof Error && w.message === "#LIMIT!"
              ? "#LIMIT!"
              : "#VALUE!",
            w instanceof Error ? w.message : "Invalid reusable formula.",
          );
        }
        l.set(x, k);
      }
      return k;
    },
    p = (x, k) => {
      let w = x.qualifier
        ? t.sheets.find(
            (N) => N.name.toLowerCase() === x.qualifier.toLowerCase(),
          )?.id
        : k;
      if (!w) return;
      let A = vn(x.name),
        I =
          t.definedNames?.find(
            (N) => N.scopeSheetId === w && N.name.toLowerCase() === A,
          ) ??
          t.definedNames?.find(
            (N) => !N.scopeSheetId && N.name.toLowerCase() === A,
          );
      return I ? { found: I, sheetId: I.scopeSheetId ?? w } : undefined;
    },
    m = (x, k, w) => {
      try {
        return o.parse(x, k, w);
      } catch (A) {
        let I =
          A && typeof A == "object" && "details" in A ? A.details : undefined;
        return o.isIssue(I) ? I : o.isIssue(A) ? A : o.issue("#ERROR!");
      }
    },
    h = (x, k, w, A, I) => {
      let N = vn(x.name);
      if (!x.qualifier && w.has(N)) return w.get(N);
      if (!x.qualifier && /^(?:TRUE|FALSE)$/i.test(x.name)) return N === "true";
      let R = p(x, A);
      if (!R)
        return c(
          "#NAME?",
          `Unknown name or unsupported custom function: ${k.slice(x.start, x.end)}.`,
        );
      let C = `${R.sheetId}:${R.found.name.toLowerCase()}`;
      if (o.allowName && !o.allowName(R.sheetId))
        return c(
          "#BLOCKED!",
          "The defined name is outside the authorized inspection context.",
        );
      if (s.has(C))
        return c("#CYCLE!", `Circular defined name: ${R.found.name}.`);
      s.add(C);
      try {
        return b(
          d(R.found.refersTo),
          R.found.refersTo,
          new Map(),
          R.sheetId,
          I + 1,
        );
      } finally {
        s.delete(C);
      }
    },
    v = (x, k, w) => {
      if (!ky(x))
        return c(
          "#VALUE!",
          "Only a LAMBDA formula can be called with arguments.",
        );
      if (k.length !== x.parameters.length)
        return c(
          "#VALUE!",
          `LAMBDA expects ${x.parameters.length} argument(s); received ${k.length}.`,
        );
      if (++i > ba.calls)
        return c("#LIMIT!", "Reusable formula evaluation exceeds 1,024 calls.");
      let A = new Map(x.environment);
      return (
        x.parameters.forEach((I, N) => A.set(I, f(k[N]))),
        b(x.body, x.source, A, x.sheetId, w + 1)
      );
    },
    g = (x, k, w, A, I) => {
      if ((u(1, I), x.kind === "symbol")) return h(x, k, w, A, I);
      if (x.kind === "group") return b(x.body, k, w, A, I + 1);
      if (x.kind === "text") return m(x.text, A, []);
      let N = gu(x);
      if (N === "LET") {
        let S = new Map(w);
        for (let F = 0; F < x.args.length - 1; F += 2) {
          let D = b(x.args[F + 1], k, S, A, I + 1),
            Q = vn(Hi(x.args[F]).name);
          (S.set(Q, f(D)), o.step?.("binding", `LET ${Q}`, D));
        }
        return b(x.args.at(-1), k, S, A, I + 1);
      }
      if (N === "LAMBDA")
        return {
          kind: "formula-closure",
          parameters: x.args.slice(0, -1).map((S) => vn(Hi(S, true).name)),
          body: x.args.at(-1),
          source: k,
          environment: new Map(w),
          sheetId: A,
        };
      let R =
          x.callee.kind === "symbol" && !x.callee.qualifier
            ? ks(x.callee.name)
            : "",
        C =
          x.callee.kind === "symbol" &&
          !x.callee.qualifier &&
          w.has(vn(x.callee.name));
      if (!C && R === "IF" && x.args.length >= 2 && x.args.length <= 3) {
        let S = b(x.args[0], k, w, A, I + 1);
        if (o.isIssue(S)) return S;
        if (!Array.isArray(S)) {
          if (typeof S == "string" && !/^(?:TRUE|FALSE)$/i.test(S))
            return o.issue("#VALUE!");
          let F = typeof S == "string" ? S.toUpperCase() === "TRUE" : !!S;
          return (
            o.step?.(
              "branch",
              `IF: ${F ? "true" : "false"} branch selected`,
              S,
            ),
            x.args[F ? 1 : 2] ? b(x.args[F ? 1 : 2], k, w, A, I + 1) : false
          );
        }
      }
      if (!C && /^(?:IFERROR|IFNA)$/.test(R) && x.args.length === 2) {
        let S = b(x.args[0], k, w, A, I + 1);
        if (!Array.isArray(S)) {
          if (
            o.isIssue(S) &&
            ["#LIMIT!", "#BLOCKED!", "#CYCLE!", "#ERROR!"].includes(S.error)
          )
            return S;
          let F = o.isIssue(S) && (R === "IFERROR" || S.error === "#N/A");
          return (
            o.step?.("branch", `${R}: ${F ? "fallback" : "value"} selected`, S),
            F ? b(x.args[1], k, w, A, I + 1) : S
          );
        }
      }
      if (!C && R && o.supported.has(R)) {
        let S = [],
          F = x.args.map((D) => y(D, k, w, A, I + 1, S));
        return m(`${R}(${F.join(",")})`, A, S);
      }
      let T = g(x.callee, k, w, A, I + 1),
        _ = x.args.map((S) => (qi(S).length ? b(S, k, w, A, I + 1) : null)),
        P = v(T, _, I);
      return (
        o.step?.("function", k.slice(x.callee.start, x.callee.end), P),
        P
      );
    },
    y = (x, k, w, A, I, N) =>
      x
        .map((R) => {
          if (R.kind === "text") {
            let T = /^(\d+(?:\.\d*)?|\.\d+)[eE]([+-]?\d+)$/.exec(R.text);
            return T ? `(${T[1]}*10^(${Number(T[2])}))` : R.text;
          }
          let C = g(R, k, w, A, I);
          return ky(C)
            ? c(
                "#CALC!",
                "A LAMBDA must be called with arguments or saved as a reusable name.",
              )
            : (f(C),
              N.push(
                Array.isArray(C)
                  ? C.map((T) => (Array.isArray(T) ? [...T] : T))
                  : C,
              ),
              `${Uf}(${N.length - 1})`);
        })
        .join(""),
    b = (x, k, w, A, I) => {
      try {
        u(x.length, I);
        let N = qi(x);
        if (N.length === 1 && N[0].kind !== "text") return g(N[0], k, w, A, I);
        let R = [];
        return m(y(x, k, w, A, I, R), A, R);
      } catch (N) {
        if (
          o.isIssue(N) &&
          !["#LIMIT!", "#BLOCKED!", "#CYCLE!", "#ERROR!"].includes(N.error)
        )
          return N;
        throw N;
      }
    };
  try {
    let x = b(d(e), e, new Map(), r, 0);
    return ky(x)
      ? c(
          "#CALC!",
          "A LAMBDA must be called with arguments or saved as a reusable name.",
        )
      : x;
  } catch (x) {
    if (o.isIssue(x)) return x;
    throw x;
  }
}
var Ay = String.raw`(?:'(?:[^']|'')+'|[^\s+\-*/^&=<>(),;:'"\[\]!]+)`,
  oB = new RegExp(`^(${Ay})!`),
  iB = new RegExp(`^(${Ay}):(${Ay})!`);
var NB = 2 * 1024 * 1024;
var P4 = ["sum", "count", "average", "min", "max"],
  oV = [...P4, "distinctCount"];
var hR = (e, t) => `${Me(t)}${e + 1}`,
  U4 = (e, t, r) =>
    t >= e.firstRow && t <= e.lastRow && r >= e.firstCol && r <= e.lastCol;
function L4(e) {
  let t = pt(e.range);
  return {
    ...t,
    firstRow: t.firstRow + 1,
    lastRow: t.lastRow - (e.showTotals ? 1 : 0),
  };
}
function Lf(e, t, r, o) {
  let n = /^([A-Za-z_][A-Za-z0-9_.]*)?(\[.+\])$/.exec(r);
  if (!n) return null;
  let i;
  if (n[1])
    for (let m of e.sheets) {
      let h = m.tables?.find(
        (v) => v.name.toLowerCase() === n[1].toLowerCase(),
      );
      if (h) {
        i = { sheet: m, table: h };
        break;
      }
    }
  else if (o) {
    let m = e.sheets.find((g) => g.id === t),
      h = je(o),
      v = m?.tables?.find((g) => U4(pt(g.range), h.row, h.col));
    m && v && (i = { sheet: m, table: v });
  }
  if (!i) return null;
  let { sheet, table } = i,
    s = pt(table.range),
    c = L4(table),
    u = n[2]
      .replace(/^\[|\]$/g, "")
      .replace(/[\[\]]/g, "")
      .split(/[:,]/)
      .map((m) => m.trim()),
    f = u.find((m) => m.startsWith("#"))?.toLowerCase(),
    d = u.some((m) => m.startsWith("@")) || f === "#this row";
  if (
    u.some(
      (m) =>
        m.startsWith("#") &&
        !["#all", "#data", "#headers", "#totals", "#this row"].includes(
          m.toLowerCase(),
        ),
    )
  )
    return null;
  if (
    (f === "#all" && Object.assign(c, s),
    f === "#headers" && Object.assign(c, s, { lastRow: s.firstRow }),
    f === "#totals")
  ) {
    if (!table.showTotals) return null;
    Object.assign(c, s, { firstRow: s.lastRow });
  }
  if (d) {
    if (!o || sheet.id !== t) return null;
    let m = je(o);
    if (m.row < c.firstRow || m.row > c.lastRow) return null;
    ((c.firstRow = m.row), (c.lastRow = m.row));
  }
  let p = u.filter((m) => !m.startsWith("#")).map((m) => m.replace(/^@/, ""));
  if (p.length) {
    if (p.length > 2) return null;
    let m = p.map((h) =>
      table.columns.findIndex((v) => v.name.toLowerCase() === h.toLowerCase()),
    );
    if (m.some((h) => h < 0)) return null;
    ((c.firstCol = s.firstCol + Math.min(...m)),
      (c.lastCol = s.firstCol + Math.max(...m)));
  }
  return c.firstRow > c.lastRow
    ? null
    : {
        sheetId: sheet.id,
        range: `${hR(c.firstRow, c.firstCol)}:${hR(c.lastRow, c.lastCol)}`,
      };
}
var vR = "\\$?[A-Za-z]{1,3}\\$?[1-9]\\d{0,6}",
  D4 = new RegExp(`^(${vR}(?::${vR})?)(?![A-Za-z0-9_.$:#\\[])`);
function Ny(e, t = 5e3) {
  let r = (p) => {
      let m = /^\$?([A-Za-z]{1,3})\$?([1-9]\d{0,6})$/.exec(p);
      if (!m)
        throw Error("External links require bounded A1 cells or rectangles.");
      let h = 0;
      for (let g of m[1].toUpperCase()) h = h * 26 + g.charCodeAt(0) - 64;
      let v = Number(m[2]);
      if (h > 16384 || v > 1048576)
        throw Error("External reference exceeds worksheet bounds.");
      return { row: v, col: h };
    },
    [o, n = o, i] = e.split(":");
  if (i) throw Error("Invalid external range.");
  let a = r(o),
    l = r(n),
    s = [],
    c = Math.min(a.row, l.row),
    u = Math.max(a.row, l.row),
    f = Math.min(a.col, l.col),
    d = Math.max(a.col, l.col);
  if ((u - c + 1) * (d - f + 1) > t)
    throw Error(
      `External link ranges are limited to ${t.toLocaleString()} cells.`,
    );
  for (let p = c; p <= u; p++)
    for (let m = f; m <= d; m++) {
      let h = m,
        v = "";
      for (; h;)
        ((v = String.fromCharCode(65 + ((h - 1) % 26)) + v),
          (h = Math.floor((h - 1) / 26)));
      s.push(v + p);
    }
  return s;
}
function F4(e) {
  let t = [];
  for (let r = 0; r < e.length;) {
    if (e[r] === '"') {
      for (r++; r < e.length;)
        if (e[r++] === '"')
          if (e[r] === '"') r++;
          else break;
      continue;
    }
    let o = r,
      n = /^'((?:[^']|'')*)'\s*!/.exec(e.slice(r)),
      i = /^\[([^\[\]\r\n]+)\]([A-Za-z_][A-Za-z0-9_.]*)!/.exec(e.slice(r)),
      a,
      l = 0;
    if (
      (n
        ? ((a = n[1].replace(/''/g, "'")), (l = n[0].length))
        : i && ((a = `[${i[1]}]${i[2]}`), (l = i[0].length)),
      a)
    ) {
      let s = /^(.*\[[^\[\]\r\n]+\])([^\[\]:\r\n]+)$/.exec(a),
        c = D4.exec(e.slice(r + l));
      if (s && c && (!r || !/[A-Za-z0-9_.\]]/.test(e[r - 1])))
        try {
          (Ny(c[1]),
            (r += l + c[0].length),
            t.push({
              start: o,
              end: r,
              token: e.slice(o, r),
              target: s[1],
              sheet: s[2],
              range: c[1].toUpperCase(),
            }));
          continue;
        } catch {}
      r += l;
      continue;
    }
    if (e[r] === "[") {
      let s = 0;
      do (e[r] === "[" && s++, e[r] === "]" && s--, r++);
      while (r < e.length && s > 0);
    } else if (e[r] === "'") {
      for (r++; r < e.length;)
        if (e[r++] === "'")
          if (e[r] === "'") r++;
          else break;
    } else r++;
  }
  return t;
}
function yR(e, t) {
  let r = F4(e);
  for (let o of r.reverse()) {
    let n = t(o);
    n !== undefined && (e = e.slice(0, o.start) + n + e.slice(o.end));
  }
  return e;
}
var fV = Object.freeze({
    sources: 8,
    sheets: 32,
    cells: 5e3,
    formulas: 500,
    metadataCharacters: 2e6,
  }),
  Yo = (e) =>
    E.string()
      .min(1)
      .max(e)
      .refine((t) => !/[\u0000-\u001f\u007f]/.test(t)),
  bR = Yo(2048).refine(
    (e) => /^[^\[\]]*\[[^\[\]]+\]$/.test(e),
    "Use an exact workbook qualifier such as [Source.xlsx].",
  ),
  wR = E.string().regex(/^[a-f0-9]{64}$/),
  z4 = E.discriminatedUnion("type", [
    E.strictObject({ type: E.literal("number"), value: E.number().finite() }),
    E.strictObject({ type: E.literal("text"), value: E.string().max(32767) }),
    E.strictObject({ type: E.literal("boolean"), value: E.boolean() }),
    E.strictObject({ type: E.literal("blank"), value: E.null() }),
    E.strictObject({
      type: E.literal("error"),
      value: E.string()
        .regex(/^#[A-Z0-9/?!_.]+$/)
        .max(30),
    }),
  ]),
  W4 = E.strictObject({
    target: bR,
    sourceId: Yo(100),
    sourceName: Yo(200),
    sourceRevision: E.number().int().nonnegative(),
    sourceFingerprint: wR,
    resolvedAt: E.iso.datetime(),
    sheets: E.array(
      E.strictObject({
        id: Yo(100),
        name: Yo(31),
        cells: E.record(E.string().regex(/^[A-Z]{1,3}[1-9]\d{0,6}$/), z4),
      }),
    )
      .min(1)
      .max(32),
  }),
  dV = E.strictObject({
    schemaVersion: E.literal(1),
    sources: E.array(W4).max(8),
    packageTargets: E.array(
      E.strictObject({
        key: bR,
        target: Yo(4096),
        part: Yo(1024),
        sourceSha256: wR,
      }),
    )
      .max(100)
      .optional(),
  });
function ER(e) {
  let t = [],
    r = new Map(),
    o = new Map(),
    n = new Set(e.sheets.map((a) => a.name.toLowerCase()));
  for (let [a, l] of (e.workbookLinks?.sources ?? []).entries())
    for (let [s, c] of l.sheets.entries()) {
      let u = `_LG_LINK_${a}_${s}`;
      for (; n.has(u.toLowerCase());) u += "_";
      n.add(u.toLowerCase());
      let f = `linked:${a}:${s}`;
      for (; e.sheets.some((p) => p.id === f);) f += "_";
      let d = {};
      for (let [p, m] of Object.entries(c.cells))
        ((d[p] = { raw: m.value, literal: true }), r.set(`${f}:${p}`, m));
      (t.push({ id: f, name: u, rowCount: 1048576, colCount: 16384, cells: d }),
        o.set(JSON.stringify([l.target, c.name.toLowerCase()]), u));
    }
  return {
    sheets: t,
    values: r,
    rewrite: (a) =>
      yR(a, (l) => {
        let s = e.workbookLinks?.sources.find((u) => u.target === l.target),
          c = s?.sheets.find(
            (u) => u.name.toLowerCase() === l.sheet.toLowerCase(),
          );
        if (!(!c || !Ny(l.range).every((u) => Object.hasOwn(c.cells, u))))
          return `'${o.get(JSON.stringify([l.target, l.sheet.toLowerCase()]))}'!${l.range}`;
      }),
  };
}
function xR(e, t, r, o, n = true) {
  let i = "";
  for (let u = 0; u < e.length;) {
    let f = u,
      d = e[u];
    if (d === '"' || d === "'") {
      for (u++; u < e.length;)
        if (e[u++] === d)
          if (e[u] === d) u++;
          else break;
      i += d === '"' ? '""' : e.slice(f, u);
      continue;
    }
    let p = /^(?:[A-Za-z_][A-Za-z0-9_.]*)?\[/.exec(e.slice(u));
    if (p) {
      u += p[0].length;
      let m = 1;
      for (; u < e.length && m;)
        (e[u] === "[" && m++, e[u] === "]" && m--, u++);
      let h = e.slice(f, u);
      i += !m && Lf(t, r, h, o) ? "TABLETOKEN" : h;
      continue;
    }
    i += e[u++];
  }
  let a = (u, f, d) => ({
    sheetId: r,
    address: o,
    kind: u,
    error: f,
    message: d,
  });
  if (
    /\[[^\]]+\][^!+*/^&=<>(),;{}]*!/.test(i) ||
    /\|/.test(i) ||
    /(?:https?:|file:|\\\\)/i.test(i)
  )
    return a(
      "external",
      "#BLOCKED!",
      "External workbook, path, network and DDE references are not supported. Use an authorized local worksheet reference.",
    );
  let l = `(?:'(?:[^']|'')*'|[^\\s!+*/^&=<>(),;:{}\\[\\]\\"'%\\-]+)`;
  if (
    /'(?:[^']|'')*:(?:[^']|'')*'\s*!/.test(i) ||
    new RegExp(`${l}\\s*:\\s*${l}\\s*!`).test(i)
  )
    return a(
      "three-dimensional",
      "#VALUE!",
      "3-D references across a worksheet interval are not supported. List the individual worksheet references explicitly.",
    );
  let c = i
    .replace(/'(?:[^']|'')*'\s*!/g, "Sheet!")
    .replace(
      /(?<![\p{L}\p{N}_.\]])\$?[A-Za-z]{1,3}\$?[1-9]\d*(?![\p{L}\p{N}_.])/gu,
      "CELLREF",
    );
  if (
    n &&
    (/(?<![\p{L}\p{N}_.])R(?:\d+|\[[+-]?\d+\])?C(?:\d+|\[[+-]?\d+\])?(?![\p{L}\p{N}_.]|\s*[!(])/iu.test(
      c,
    ) ||
      /(?<![\p{L}\p{N}_.])[RC]\[[+-]?\d+\](?![\p{L}\p{N}_.])/iu.test(c))
  )
    return a(
      "r1c1",
      "#REF!",
      "R1C1-style references are not supported. Use A1 references, including $A$1, $A1 or A$1 as needed.",
    );
}
var Ry = [
    "DSUM",
    "DAVERAGE",
    "DCOUNT",
    "DCOUNTA",
    "DMAX",
    "DMIN",
    "DPRODUCT",
    "DSTDEV",
    "DSTDEVP",
    "DVAR",
    "DVARP",
    "DGET",
  ],
  Ty = Symbol("not a database function"),
  vo = {
    databaseCells: 2e4,
    criteriaCells: 1e3,
    predicateComparisons: 1e5,
    textComparisonSteps: 5e6,
    criterionCharacters: 255,
    headerCharacters: 255,
    formulaEvaluations: 1e4,
  };
function IR(e, t, r) {
  if (!Ry.includes(e)) return Ty;
  let o = (D = "#VALUE!") => {
    throw r.issue(D);
  };
  if (t.length !== 3 || t[0] == null || t[2] == null) return o();
  let n = (D, Q) => {
      let te = r.retrieve(D);
      if (r.isIssue(te)) throw te;
      if (!Array.isArray(te) || te.length < 2) return o();
      if (te.length > Q) return o("#LIMIT!");
      if (!Array.isArray(te[0])) return o();
      let K = te,
        Z = K[0].length;
      if (!Z) return o();
      if (K.length * Z > Q) return o("#LIMIT!");
      for (let B = 0; B < K.length; B++)
        if (!Array.isArray(K[B]) || K[B].length !== Z) return o();
      for (let B of K)
        for (let X of B)
          if (r.isIssue(X)) {
            if (
              [
                "#LIMIT!",
                "#CYCLE!",
                "#BLOCKED!",
                "#TIMEOUT!",
                "#ITER!",
                "#ERROR!",
              ].includes(X.error)
            )
              throw X;
          } else {
            if (
              X != null &&
              !["number", "string", "boolean"].includes(typeof X)
            )
              return o();
            if (typeof X == "number" && !Number.isFinite(X)) return o("#NUM!");
          }
      return K;
    },
    i = n(t[0], vo.databaseCells),
    a = n(t[2], vo.criteriaCells);
  if ((i.length - 1) * (a.length - 1) * a[0].length > vo.predicateComparisons)
    return o("#LIMIT!");
  let l = (D) => {
      if (r.isIssue(D)) throw D;
      return typeof D != "string" || !D.trim() || D.length > vo.headerCharacters
        ? o()
        : D.toLowerCase();
    },
    s = i[0].map(l);
  if (new Set(s).size !== s.length) return o();
  let c = a[0].map((D) => {
      if (r.isIssue(D)) throw D;
      return D == null || D === "" ? -1 : s.indexOf(l(D));
    }),
    u = e === "DCOUNT" || e === "DCOUNTA",
    f;
  if (t[1] == null) {
    if (!u) return o();
  } else {
    let D = r.retrieve(t[1]);
    if (r.isIssue(D)) throw D;
    if (typeof D == "number") {
      if (!Number.isInteger(D) || D < 1 || D > s.length) return o();
      f = D - 1;
    } else if (typeof D == "string" && D.length) {
      if (((f = s.indexOf(D.toLowerCase())), f < 0)) return o();
    } else return o();
  }
  let d = 0,
    p = (D) => {
      ((d += D), d > vo.textComparisonSteps && o("#LIMIT!"));
    },
    m = (D, Q, te) =>
      te === "<>"
        ? D !== Q
        : te === "<"
          ? D < Q
          : te === ">"
            ? D > Q
            : te === "<="
              ? D <= Q
              : te === ">="
                ? D >= Q
                : D === Q,
    h = (D) => {
      if (r.isIssue(D)) throw D;
      if (D == null || D === "") return null;
      if (typeof D == "number" || typeof D == "boolean") return (L) => L === D;
      if (typeof D != "string") return o();
      if (D.length > vo.criterionCharacters) return o("#LIMIT!");
      let Q = /^(<=|>=|<>|=|<|>)?(.*)$/s.exec(D),
        te = Q[1] ?? "=",
        K = Q[2];
      if (Q[1] && /^[<>=]/.test(K)) return o();
      if (!K && Q[1])
        return te !== "=" && te !== "<>"
          ? o()
          : (L) => (te === "=" ? L == null || L === "" : L != null && L !== "");
      if (/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(K.trim())) {
        let L = Number(K);
        return Number.isFinite(L)
          ? (U) => (typeof U == "number" ? m(U, L, te) : te === "<>")
          : o("#NUM!");
      }
      if (/^(TRUE|FALSE)$/i.test(K)) {
        if (te !== "=" && te !== "<>") return o();
        let L = K.toUpperCase() === "TRUE";
        return (U) => (te === "=" ? U === L : U !== L);
      }
      let Z = K.toLowerCase(),
        B = [],
        X = false;
      for (let L = 0; L < Z.length; L++) {
        let U = Z[L];
        if (U === "~") {
          let Y = Z[L + 1];
          Y && ["*", "?", "~"].includes(Y)
            ? (B.push({ kind: "literal", char: Y }), L++, (X = true))
            : B.push({ kind: "literal", char: U });
        } else
          U === "*" || U === "?"
            ? ((X = true),
              (U === "?" || B.at(-1)?.kind !== "many") &&
                B.push({ kind: U === "*" ? "many" : "one" }))
            : B.push({ kind: "literal", char: U });
      }
      return ["<", ">", "<=", ">="].includes(te)
        ? X
          ? o()
          : (L) =>
              typeof L != "string"
                ? false
                : (p(L.length + Z.length), m(L.toLowerCase(), Z, te))
        : (!Q[1] && !X && B.push({ kind: "many" }),
          (L) => {
            if (typeof L != "string") return te === "<>";
            p((L.length + 1) * (B.length + 1));
            let U = L.toLowerCase(),
              Y = new Uint8Array(U.length + 1);
            Y[0] = 1;
            for (let ve of B) {
              let $ = new Uint8Array(U.length + 1);
              if (ve.kind === "many") {
                $[0] = Y[0];
                for (let W = 1; W <= U.length; W++) $[W] = Y[W] || $[W - 1];
              } else
                for (let W = 1; W <= U.length; W++)
                  $[W] =
                    Y[W - 1] && (ve.kind === "one" || ve.char === U[W - 1])
                      ? 1
                      : 0;
              Y = $;
            }
            let ie = !!Y[U.length];
            return te === "<>" ? !ie : ie;
          });
    },
    v = 0,
    g = a
      .slice(1)
      .map((D, Q) =>
        D.map((te, K) =>
          c[K] < 0 ? r.criterionFormula?.(t[2], Q + 1, K) : undefined,
        ),
      );
  c.forEach((D, Q) => {
    D < 0 && !g.some((te) => te[Q]) && o();
  });
  let y = a.slice(1).map((D, Q) =>
    D.map((te, K) => {
      if (c[K] >= 0) return h(te);
      if (!g[Q][K] && !(te == null || te === "")) return o();
    }),
  );
  g.flat().filter(Boolean).length * (i.length - 1) > vo.formulaEvaluations &&
    o("#LIMIT!");
  let x = 0,
    k = 0,
    w = 0,
    A = 0,
    I = 1,
    N = 1 / 0,
    R = -1 / 0,
    C = 0,
    T = 0,
    _;
  for (let [D, Q] of i.slice(1).entries()) {
    let te = false;
    for (let [B, X] of y.entries()) {
      let L = true;
      for (let U = 0; U < X.length; U++) {
        let Y = g[B][U];
        if (Y) {
          ++v > vo.formulaEvaluations && o("#LIMIT!");
          let $ = Y(D);
          if (r.isIssue($)) throw $;
          (typeof $ != "boolean" && o(), $ || (L = false));
          continue;
        }
        let ie = X[U];
        if (!ie) continue;
        let ve = Q[c[U]];
        if (r.isIssue(ve)) throw ve;
        ie(ve) || (L = false);
      }
      L && (te = true);
    }
    if (!te || (x++, f === undefined)) continue;
    let K = Q[f];
    if (((_ = K), K != null && w++, r.isIssue(K))) {
      if (!["DCOUNT", "DCOUNTA", "DGET"].includes(e)) throw K;
      continue;
    }
    if (typeof K != "number") continue;
    (k++, (A += K), (I *= K), (N = Math.min(N, K)), (R = Math.max(R, K)));
    let Z = K - C;
    ((C += Z / k), (T += Z * (K - C)));
  }
  if (e === "DGET") {
    if (x !== 1) return o(x ? "#NUM!" : "#VALUE!");
    if (r.isIssue(_)) throw _;
    return _ ?? 0;
  }
  if (e === "DCOUNT") return f === undefined ? x : k;
  if (e === "DCOUNTA") return f === undefined ? x : w;
  let P = (D) => (Number.isFinite(D) ? D : o("#NUM!"));
  if (e === "DSUM") return P(A);
  if (e === "DMAX") return k ? R : 0;
  if (e === "DMIN") return k ? N : 0;
  if (e === "DPRODUCT") return k ? P(I) : 0;
  if (e === "DAVERAGE") return k ? P(A / k) : o("#DIV/0!");
  let S = e === "DSTDEV" || e === "DVAR";
  if (k < (S ? 2 : 1)) return o("#DIV/0!");
  let F = Math.max(0, T / (k - (S ? 1 : 0)));
  return P(e === "DSTDEV" || e === "DSTDEVP" ? Math.sqrt(F) : F);
}
var Cy = ["SUBTOTAL", "AGGREGATE", "ISFORMULA", "FORMULATEXT"],
  Sy = Symbol("not-reference-function"),
  yo = (e) =>
    !!e &&
    typeof e == "object" &&
    !Array.isArray(e) &&
    [Object.prototype, null].includes(Object.getPrototypeOf(e)),
  Df = (e) => typeof e?.raw == "string" && e.raw.startsWith("=") && !e.literal;
function B4(e) {
  return /(?:^|[^A-Za-z0-9_.])(?:_xlfn\.(?:_xlws\.)?)?SUBTOTAL\s*\(/i.test(
    e.replace(/"(?:[^"]|"")*"|'(?:[^']|'')*'(?=\s*!)/g, ""),
  );
}
function V4(e) {
  return /(?:^|[^A-Za-z0-9_.])(?:_xlfn\.(?:_xlws\.)?)?(?:SUBTOTAL|AGGREGATE)\s*\(/i.test(
    e.replace(/"(?:[^"]|"")*"|'(?:[^']|'')*'(?=\s*!)/g, ""),
  );
}
function kR(e, t = {}) {
  if (!yo(t) || Object.keys(t).some((s) => s !== "filteredRowsBySheet"))
    throw Error("Calculation options support only filteredRowsBySheet.");
  let r = t.filteredRowsBySheet;
  if (r !== undefined && !yo(r))
    throw Error(
      "Filtered rows must be a worksheet-ID map of inclusive row intervals.",
    );
  let o = new Map(e.sheets.map((s) => [s.id, s])),
    n = new Map(e.sheets.map((s) => [s.name.toLowerCase(), s])),
    i = new Map(),
    a = 0;
  for (let [s, c] of Object.entries(r ?? {})) {
    let u = o.get(s);
    if (!u) throw Error("Filter context refers to an unknown worksheet ID.");
    if (!Array.isArray(c) || (a += c.length) > 1e4)
      throw Error(
        "Filter context supports at most 10,000 intervals across the workbook.",
      );
    i.set(s, Ss(u.rowCount, Vo(c, u.rowCount)));
  }
  let l = new Map();
  for (let s of e.sheets) l.set(s.id, Ss(s.rowCount, ea(s, "row")));
  return (s, c, u) => {
    if (!Cy.includes(s)) return Sy;
    let { issue, isIssue } = u,
      p = (L) => {
        if (isIssue(L)) throw L;
        if (!yo(L) || !yo(L.ref)) throw issue("#VALUE!");
        let U = L.ref;
        if (U.sheet !== undefined && typeof U.sheet != "string")
          throw issue("#REF!");
        let Y = n.get(
          String(U.sheet ?? u.currentSheetName ?? "").toLowerCase(),
        );
        if (!Y) throw issue("#REF!");
        let ie = yo(U.from) ? U.from : U,
          ve = yo(U.to) ? U.to : U,
          $ = [ie.row, ie.col, ve.row, ve.col];
        if ($.some((ue) => !Number.isInteger(ue))) throw issue("#REF!");
        let [W, se, me, we] = $;
        if (
          W < 1 ||
          se < 1 ||
          me < W ||
          we < se ||
          me > Y.rowCount ||
          we > Y.colCount
        )
          throw issue("#REF!");
        return {
          sheet: Y,
          firstRow: W,
          lastRow: me,
          firstCol: se,
          lastCol: we,
        };
      };
    if (s === "FORMULATEXT" || s === "ISFORMULA") {
      if (c.length !== 1) return issue("#VALUE!");
      let L = p(c[0]),
        U = (ie, ve) => {
          let $ = `${Me(ve - 1)}${ie}`;
          return (u.readMetadata?.(L.sheet.id, $), L.sheet.cells[$]);
        };
      if (s === "FORMULATEXT") {
        let ie = U(L.firstRow, L.firstCol);
        return Df(ie) && ie.raw.length <= 8192 ? ie.raw : issue("#N/A");
      }
      let Y = (L.lastRow - L.firstRow + 1) * (L.lastCol - L.firstCol + 1);
      return Y > 1e4
        ? issue("#LIMIT!")
        : Y === 1
          ? Df(U(L.firstRow, L.firstCol))
          : Array.from({ length: L.lastRow - L.firstRow + 1 }, (ie, ve) =>
              Array.from({ length: L.lastCol - L.firstCol + 1 }, ($, W) =>
                Df(U(L.firstRow + ve, L.firstCol + W)),
              ),
            );
    }
    let m = s === "AGGREGATE";
    if (c.length < (m ? 3 : 2) || c.length > 255) return issue("#VALUE!");
    let h = (L) => {
        let U = u.retrieve(L);
        if (isIssue(U)) throw U;
        if (
          Array.isArray(U) ||
          !["number", "string", "boolean"].includes(typeof U) ||
          (typeof U == "string" && U.trim() === "") ||
          !Number.isFinite(Number(U))
        )
          throw issue("#VALUE!");
        return Number(U);
      },
      v = h(c[0]);
    if (m && !Number.isInteger(v)) return issue("#VALUE!");
    let g = Math.trunc(v),
      y = !m && g > 100 ? g - 100 : g;
    if (y < 1 || y > (m ? 19 : 11)) return issue("#VALUE!");
    let b = m ? (c[1] == null ? 0 : h(c[1])) : 0;
    if (!Number.isInteger(b) || b < 0 || b > 7) return issue("#VALUE!");
    let x = m ? b % 2 === 1 : g > 100,
      k = m && [2, 3, 6, 7].includes(b),
      w = !m || b < 4,
      A = m && y >= 14;
    if (A && (c.length !== 4 || c[3] == null)) return issue("#VALUE!");
    let I = A ? h(c[3]) : 0,
      N = A ? [c[2]] : c.slice(m ? 2 : 1),
      C = A && !((L) => yo(L) && yo(L.ref))(N[0]),
      T = C ? [] : N.map(p),
      _ = 0;
    for (let L of T) {
      let U = (L.lastRow - L.firstRow + 1) * (L.lastCol - L.firstCol + 1);
      if (((_ += U), U > 2e4 || _ > 1e5)) return issue("#LIMIT!");
    }
    for (let L of T)
      for (let U = L.firstRow; U <= L.lastRow; U++)
        for (let Y = L.firstCol; Y <= L.lastCol; Y++)
          if (u.isActive(L.sheet.id, `${Me(Y - 1)}${U}`))
            return issue("#CYCLE!");
    let P = 0,
      S = 0,
      F = 1 / 0,
      D = -1 / 0,
      Q = 0,
      te = 0,
      K = [],
      Z = (L) => {
        if (isIssue(L)) {
          if (["#LIMIT!", "#CYCLE!"].includes(L.error)) return L;
          if (k) return;
          if (y !== 2 && y !== 3) return L;
        }
        if ((L != null && S++, typeof L != "number")) return;
        if (!Number.isFinite(L)) return issue("#NUM!");
        (P++, (F = Math.min(F, L)), (D = Math.max(D, L)));
        let U = L - Q;
        ((Q += U / P),
          (te += U * (L - Q)),
          (y === 1 || y === 6 || y === 9 || y >= 12) && K.push(L));
      };
    if (C) {
      let L = u.retrieve(N[0]),
        U = Array.isArray(L) ? L : [[L]];
      if (!U.length || U.length > 1e4) return issue("#LIMIT!");
      let Y = U.every(Array.isArray) ? U : [U],
        ie = Y[0]?.length ?? 0;
      if (!ie || Y.some((ve) => ve.length !== ie || ve.some(Array.isArray)))
        return issue("#VALUE!");
      if (Y.length * ie > 1e4) return issue("#LIMIT!");
      for (let ve of Y)
        for (let $ of ve) {
          if (
            $ != null &&
            !isIssue($) &&
            !["number", "string", "boolean"].includes(typeof $)
          )
            return issue("#VALUE!");
          let W = Z($);
          if (W) return W;
        }
    }
    for (let L of T)
      for (let U = L.firstRow; U <= L.lastRow; U++)
        if (!(
          i.get(L.sheet.id)?.isHidden(U - 1) ||
          (x && l.get(L.sheet.id).isHidden(U - 1))
        ))
          for (let Y = L.firstCol; Y <= L.lastCol; Y++) {
            let ie = `${Me(Y - 1)}${U}`;
            u.readMetadata?.(L.sheet.id, ie);
            let ve = L.sheet.cells[ie];
            if (w && Df(ve) && (m ? V4 : B4)(u.expandedFormula(ve, L.sheet.id)))
              continue;
            let $ = Z(u.readCell(L.sheet.id, ie));
            if ($) return $;
          }
    if (y >= 12) {
      if ((K.sort((W, se) => W - se), y === 13)) {
        let W = 1,
          se = 0,
          me = 0,
          we;
        for (let ue of K)
          ((me = ue === we ? me + 1 : 1),
            me > W && ((W = me), (se = ue)),
            (we = ue));
        return W > 1 ? se : issue("#N/A");
      }
      if (!K.length) return issue("#NUM!");
      if (y === 12) {
        let W = Math.floor(K.length / 2);
        return K.length % 2 ? K[W] : K[W - 1] / 2 + K[W] / 2;
      }
      if (y === 14 || y === 15) {
        let W = I;
        return Number.isInteger(W) && W >= 1 && W <= K.length
          ? K[y === 14 ? K.length - W : W - 1]
          : issue("#NUM!");
      }
      let L = y >= 18,
        U = y === 17 || y === 19 ? Math.trunc(I) / 4 : I;
      if (U < 0 || U > 1 || (L && (U === 0 || U === 1))) return issue("#NUM!");
      let Y = L ? U * (K.length + 1) - 1 : U * (K.length - 1);
      if (Y < 0 || Y > K.length - 1) return issue("#NUM!");
      let ie = Math.floor(Y),
        ve = Math.ceil(Y),
        $ = Y - ie;
      return K[ie] * (1 - $) + K[ve] * $;
    }
    if (y === 2) return P;
    if (y === 3) return S;
    if (y === 9) return Nr(K);
    if (y === 1) return P ? Zt(K) : issue("#DIV/0!");
    if (y === 4) return P ? D : 0;
    if (y === 5) return P ? F : 0;
    if (y === 6) return P ? _f(K) : 0;
    let B = y === 7 || y === 10;
    if (P < (B ? 2 : 1)) return issue("#DIV/0!");
    let X = Math.max(0, te / (P - (B ? 1 : 0)));
    return y === 7 || y === 8 ? Math.sqrt(X) : X;
  };
}
var j4 = [
    "SUM",
    "AVERAGE",
    "MIN",
    "MAX",
    "COUNT",
    "COUNTA",
    "COUNTBLANK",
    "SUMIF",
    "SUMIFS",
    "COUNTIF",
    "COUNTIFS",
    "AVERAGEIF",
    "AVERAGEIFS",
    "IF",
    "IFERROR",
    "IFNA",
    "AND",
    "OR",
    "NOT",
    "TRUE",
    "FALSE",
    "ABS",
    "ROUND",
    "ROUNDDOWN",
    "ROUNDUP",
    "INT",
    "MOD",
    "SIGN",
    "SQRT",
    "POWER",
    "SUMPRODUCT",
    "VLOOKUP",
    "HLOOKUP",
    "INDEX",
    "MATCH",
    "LEFT",
    "RIGHT",
    "MID",
    "LEN",
    "TRIM",
    "CLEAN",
    "UPPER",
    "LOWER",
    "CONCATENATE",
    "TEXT",
    "VALUE",
    "SUBSTITUTE",
    "REPLACE",
    "FIND",
    "SEARCH",
    "DATE",
    "DAY",
    "MONTH",
    "YEAR",
    "TODAY",
    "EDATE",
    "EOMONTH",
    "DAYS",
    "NETWORKDAYS",
    "WORKDAY",
    "YEARFRAC",
    "ISBLANK",
    "ISNUMBER",
    "ISTEXT",
    "ISERROR",
    "ISNA",
    "ROW",
    "COLUMN",
    "ROWS",
    "COLUMNS",
    "LARGE",
    "SMALL",
    "MEDIAN",
    "STDEV.S",
    "VAR.S",
  ],
  Pf = [
    ...new Set([
      ...j4,
      ...vu,
      ...Ey,
      ...P0,
      ...Cy,
      ...Ry,
      ...vy,
      ...oR,
      ...dR,
      ...py,
      ...uy,
      ...dy,
    ]),
  ].sort(),
  Z4 = new Set(Pf),
  AR = $y.default.FormulaError,
  Oy = (e) => new AR(e),
  G4 = Oy,
  Ce = (e) => e instanceof AR,
  nr = (e) => (Array.isArray(e) ? e.flatMap(nr) : [e]),
  _y = (e) => e.flatMap(nr).find(Ce);
function al(e, t) {
  if (typeof t != "string") return e === t;
  let r = /^(<=|>=|<>|=|<|>)?(.*)$/.exec(t),
    o = r[1] ?? "=",
    n = r[2];
  if ((o === "=" || o === "<>") && /[?*~]/.test(n)) {
    if (typeof e != "string") return o === "<>";
    let s = [];
    for (let f = 0; f < n.length; f++)
      n[f] === "~" && f + 1 < n.length
        ? s.push({ char: n[++f].toLowerCase(), wildcard: false })
        : s.push({
            char: n[f].toLowerCase(),
            wildcard: n[f] === "*" || n[f] === "?",
          });
    let c = String(e ?? "").toLowerCase();
    if (s.length * c.length > 1e5) throw G4("#LIMIT!");
    let u = Array(c.length + 1).fill(false);
    u[0] = true;
    for (let f of s) {
      let d = Array(c.length + 1).fill(false);
      d[0] = f.wildcard && f.char === "*" && u[0];
      for (let p = 1; p <= c.length; p++)
        d[p] =
          f.wildcard && f.char === "*"
            ? u[p] || d[p - 1]
            : u[p - 1] &&
              ((f.wildcard && f.char === "?") || c[p - 1] === f.char);
      u = d;
    }
    return o === "<>" ? !u[c.length] : u[c.length];
  }
  let i = n.trim() !== "" && Number.isFinite(Number(n)),
    a = i
      ? e === null || e === ""
        ? 0
        : typeof e == "number"
          ? e
          : NaN
      : String(e ?? "").toLowerCase(),
    l = i ? Number(n) : n.toLowerCase();
  return o === "<>"
    ? a !== l
    : o === "<"
      ? a < l
      : o === ">"
        ? a > l
        : o === "<="
          ? a <= l
          : o === ">="
            ? a >= l
            : a === l;
}
function q4(e, t, r = Oy, o = Ts(), n, i) {
  let a = e.utils._applyInfix.bind(e.utils);
  ((e.utils._applyInfix = (s, c, u) => {
    let f = (g, y) => (
      ["=", "<>", "<", ">", "<=", ">="].includes(c) &&
        typeof g == "string" &&
        typeof y == "string" &&
        ((g = g.toLowerCase()), (y = y.toLowerCase())),
      a({ val: g, isArray: false }, c, { val: y, isArray: false })
    );
    if (!Array.isArray(s.val) && !Array.isArray(u.val)) return f(s.val, u.val);
    let d = (g) =>
        Array.isArray(g) ? (g.every(Array.isArray) ? g : [g]) : [[g]],
      p = d(s.val),
      m = d(u.val),
      h = Math.max(p.length, m.length),
      v = Math.max(p[0]?.length ?? 0, m[0]?.length ?? 0);
    return !h ||
      !v ||
      (p.length !== h && p.length !== 1) ||
      (m.length !== h && m.length !== 1) ||
      (p[0].length !== v && p[0].length !== 1) ||
      (m[0].length !== v && m[0].length !== 1)
      ? r("#VALUE!")
      : h * v > en.maxSpillCells
        ? r("#LIMIT!")
        : Array.from({ length: h }, (g, y) =>
            Array.from({ length: v }, (b, x) =>
              f(
                p[p.length === 1 ? 0 : y][p[0].length === 1 ? 0 : x],
                m[m.length === 1 ? 0 : y][m[0].length === 1 ? 0 : x],
              ),
            ),
          );
  }),
    (e.checkFormulaResult = (s, c) => {
      let u = e.retrieveRef(s);
      return Ce(u)
        ? u
        : Array.isArray(u)
          ? c
            ? u
            : r("#VALUE!")
          : u === null
            ? 0
            : typeof u == "number"
              ? Number.isFinite(u)
                ? u
                : r("#NUM!")
              : typeof u == "string" || typeof u == "boolean"
                ? u
                : r("#VALUE!");
    }));
  let l = e._callFunction.bind(e);
  return (
    (e._callFunction = (s, c) => {
      s = s
        .replace(/^_xlfn\./i, "")
        .replace(/^_xlws\./i, "")
        .toUpperCase();
      let u = (f) => (f == null ? undefined : e.retrieveRef(f));
      try {
        if (!Z4.has(s)) return r("#NAME?");
        if (s === "RAND") return c.length ? r("#VALUE!") : Math.random();
        if (s === "RANDBETWEEN") {
          if (c.length !== 2) return r("#VALUE!");
          let T = c.map(u),
            _ = _y(T);
          if (_) return _;
          if (
            T.some(
              (F) =>
                Array.isArray(F) ||
                F === undefined ||
                (typeof F == "string" && F.trim() === "") ||
                !Number.isFinite(Number(F)),
            )
          )
            return r("#VALUE!");
          let P = Math.ceil(Number(T[0])),
            S = Math.floor(Number(T[1]));
          return !Number.isSafeInteger(P) ||
            !Number.isSafeInteger(S) ||
            P > S ||
            !Number.isSafeInteger(S - P + 1)
            ? r("#NUM!")
            : P + Math.floor(Math.random() * (S - P + 1));
        }
        let f = Cs(s, c, { retrieve: u, issue: r, isIssue: Ce, clock: o });
        if (f !== O0) return f;
        let d = ux(s, c, {
          retrieve: u,
          issue: r,
          isIssue: Ce,
          clock: o,
          report: n,
        });
        if (d !== U0) return d;
        let p = t(s, c, e);
        if (p !== Sy) return p;
        let m = IR(s, c, {
          retrieve: u,
          issue: r,
          isIssue: Ce,
          criterionFormula: i,
        });
        if (m !== Ty) return m;
        let h = rR(s, c, { retrieve: u, issue: r, isIssue: Ce });
        if (h !== hy) return h;
        let v = YN(s, c, { retrieve: u, issue: r, isIssue: Ce });
        if (v !== fy) return v;
        let g = QN(s, c, {
          retrieve: u,
          issue: r,
          isIssue: Ce,
          distribution: l,
        });
        if (g !== my) return g;
        let y = JN(s, c, { retrieve: u, issue: r, isIssue: Ce });
        if (y !== cy) return y;
        let b = nR(s, c, { retrieve: u, issue: r, isIssue: Ce });
        if (b !== yy) return b;
        let x = aR(s, c, { retrieve: u, issue: r, isIssue: Ce });
        if (x !== wy) return x;
        let k = iR(s, c, { retrieve: u, issue: r, isIssue: Ce, matches: al });
        if (k !== by) return k;
        if (s === "TEXTBEFORE" || s === "TEXTAFTER")
          return Ru(
            c.map(u),
            (T) => {
              try {
                return D0(s, T, {
                  retrieve: (_) => _,
                  issue: r,
                  isIssue: Ce,
                  format: () => r("#VALUE!"),
                });
              } catch (_) {
                return Ce(_) ? _ : r("#VALUE!");
              }
            },
            r,
          );
        let w = D0(s, c, {
          retrieve: u,
          issue: r,
          isIssue: Ce,
          temporal: (T) => Un(T, o),
          format: () => {
            let T = Number(u(c[0]) ?? 0),
              _ = u(c[1]);
            if (typeof T == "number" && typeof _ == "string") {
              try {
                if (Os(_).kind === "temporal")
                  return hx(
                    T,
                    { category: "custom", locale: "en-US", pattern: _ },
                    o.dateSystem,
                  ).text;
              } catch {}
              let P = _.replace(/"[^"]*"|\\./g, "");
              if (
                o.dateSystem === "1904" &&
                /[ydm]/i.test(P) &&
                !/\[[hms]+\]/i.test(P)
              )
                return l(s, [T + Rs(o.dateSystem), _]);
            }
            return l(s, c);
          },
        });
        if (w !== Tu) return w;
        let A = cR(s, c, { retrieve: u, issue: r, isIssue: Ce, matches: al });
        if (A !== Mf) return A;
        if (s === "INDEX") {
          if (c.length < 2 || c.length > 4) return r("#VALUE!");
          if (Ce(c[0])) return c[0];
          let T = [...c];
          for (let P = 1; P < c.length; P++) {
            let S = c[P] == null ? (P === 3 ? 1 : 0) : u(c[P]);
            if (Ce(S)) return S;
            if (
              Array.isArray(S) ||
              (typeof S == "string" && S.trim() === "") ||
              !Number.isFinite(Number(S)) ||
              Number(S) < 0
            )
              return r("#VALUE!");
            T[P] = Number(S);
          }
          if (T.length > 3 && Number(T[3]) < 1) return r("#VALUE!");
          let _ = l(s, T);
          if (_ && typeof _ == "object" && "ref" in _) {
            let P = c[0],
              S = _;
            P?.ref?.sheet && !S.ref.sheet && (S.ref.sheet = P.ref.sheet);
          }
          return _;
        }
        if (s === "IF") {
          if (c.length < 2 || c.length > 3) return r("#VALUE!");
          let T = u(c[0]);
          if (Ce(T)) return T;
          if (Array.isArray(T))
            return Ru(
              [T, u(c[1]), c.length === 3 ? u(c[2]) : false],
              ([P, S, F]) =>
                Ce(P)
                  ? P
                  : typeof P == "string" && !/^(TRUE|FALSE)$/i.test(P)
                    ? r("#VALUE!")
                    : (((
                        typeof P == "string" ? P.toUpperCase() === "TRUE" : !!P
                      )
                        ? S
                        : F) ?? 0),
              r,
            );
          if (typeof T == "string" && !/^(TRUE|FALSE)$/i.test(T))
            return r("#VALUE!");
          let _ = typeof T == "string" ? T.toUpperCase() === "TRUE" : !!T;
          return u(_ ? c[1] : c.length === 3 ? c[2] : false) ?? 0;
        }
        if (s === "IFERROR" || s === "IFNA") {
          if (c.length !== 2) return r("#VALUE!");
          let T = u(c[0]);
          return Array.isArray(T)
            ? Ru(
                [T, u(c[1])],
                ([_, P]) =>
                  Ce(_) && (s === "IFERROR" || _.error === "#N/A")
                    ? (P ?? 0)
                    : (_ ?? 0),
                r,
              )
            : Ce(T) && (s === "IFERROR" || T.error === "#N/A")
              ? (u(c[1]) ?? 0)
              : (T ?? 0);
        }
        let I = c.map(u),
          N = I.flatMap(nr).filter((T) => typeof T == "number");
        if (
          I.flatMap(nr).length > 1e5 ||
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
          ].includes(s) &&
            ((Math.abs(Number(I[s === "BASE" ? 2 : 1] ?? 0)) > 32767 &&
              !["FIXED", "DOLLAR"].includes(s)) ||
              (["FIXED", "DOLLAR"].includes(s) &&
                Math.abs(Number(I[1] ?? 2)) > 100) ||
              (s === "REPT" &&
                String(I[0] ?? "").length * Number(I[1]) > 32767))) ||
          (["MUNIT", "MDETERM", "MMULT", "GROWTH"].includes(s) &&
            (s === "MUNIT"
              ? Number(I[0]) > 100
              : I.flatMap(nr).length > 1e3)) ||
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
          ].includes(s) &&
            N.some((T) => Math.abs(T) > 1e4)) ||
          (["BESSELI", "BESSELJ", "BESSELK", "BESSELY"].includes(s) &&
            Math.abs(Number(I[1])) > 1e3)
        )
          return r("#LIMIT!");
        if (s === "ROMAN" && Math.abs(Number(I[0])) > 3999) return r("#VALUE!");
        if (
          [
            "NETWORKDAYS",
            "NETWORKDAYS.INTL",
            "WORKDAY",
            "WORKDAY.INTL",
          ].includes(s) &&
          (Math.abs(
            Number(I[1]) - (s.startsWith("NETWORK") ? Number(I[0]) : 0),
          ) > 366e3 ||
            I.flatMap(nr).length > 1e3)
        )
          return r("#LIMIT!");
        if (s === "ISERR") return Ce(I[0]) && I[0].error !== "#N/A";
        if (s === "ISLOGICAL") return typeof I[0] == "boolean";
        if (s === "ISNONTEXT") return typeof I[0] != "string";
        if (s === "N")
          return Ce(I[0]) || typeof I[0] == "number"
            ? I[0]
            : I[0] === true
              ? 1
              : 0;
        if (s === "T") return Ce(I[0]) || typeof I[0] == "string" ? I[0] : "";
        if (s === "TYPE")
          return Ce(I[0])
            ? 16
            : Array.isArray(I[0])
              ? 64
              : typeof I[0] == "string"
                ? 2
                : typeof I[0] == "boolean"
                  ? 4
                  : 1;
        if (s === "ERROR.TYPE") {
          let T = [
            "#NULL!",
            "#DIV/0!",
            "#VALUE!",
            "#REF!",
            "#NAME?",
            "#NUM!",
            "#N/A",
            "#GETTING_DATA",
          ].indexOf(Ce(I[0]) ? I[0].error : "");
          return T < 0 ? r("#N/A") : T + 1;
        }
        if (s === "ISERROR") return Ce(I[0]);
        if (s === "ISNA") return Ce(I[0]) && I[0].error === "#N/A";
        if (s === "ISNUMBER") return typeof I[0] == "number";
        if (s === "ISTEXT") return typeof I[0] == "string";
        if (s === "ISBLANK") return I[0] === null;
        if (s === "COUNTA")
          return I.flatMap(nr).filter((T) => T !== null).length;
        if (s === "COUNTBLANK")
          return I.flatMap(nr).filter((T) => T === null || T === "").length;
        if (
          [
            "SUMIF",
            "AVERAGEIF",
            "COUNTIF",
            "SUMIFS",
            "AVERAGEIFS",
            "COUNTIFS",
          ].includes(s)
        ) {
          let T = s.endsWith("IFS"),
            _ = s.startsWith("COUNT"),
            P = s.startsWith("AVERAGE"),
            S,
            F;
          if (T) {
            if (((F = _ ? I : I.slice(1)), F.length < 2 || F.length % 2 !== 0))
              return r("#VALUE!");
            S = _ ? nr(F[0]).map(() => 1) : nr(I[0]);
          } else {
            if (I.length < 2 || I.length > (_ ? 2 : 3)) return r("#VALUE!");
            ((F = I.slice(0, 2)),
              (S = _
                ? nr(I[0]).map(() => 1)
                : nr(c.length < 3 || c[2] == null ? I[0] : I[2])));
          }
          let D = [],
            Q = [],
            te = (X) => {
              if (!Array.isArray(X)) return "1:1";
              let L = X.every(Array.isArray) ? X : [X],
                U = L[0]?.length ?? 0;
              return L.length && U && L.every((Y) => Y.length === U)
                ? `${L.length}:${U}`
                : "invalid";
            },
            K = te(_ ? F[0] : T || c.length < 3 || c[2] == null ? I[0] : I[2]);
          for (let X = 0; X < F.length; X += 2) {
            let L = nr(F[X]);
            if (
              L.length !== S.length ||
              K === "invalid" ||
              te(F[X]) !== K ||
              Array.isArray(F[X + 1])
            )
              return r("#VALUE!");
            if (Ce(F[X + 1])) return F[X + 1];
            D.push(L);
            let U = F[X + 1];
            Q.push(U == null ? "=0" : typeof U == "number" ? `=${U}` : U);
          }
          let Z = [],
            B = 0;
          for (let X = 0; X < S.length; X++) {
            if (
              ((B += D.reduce(
                (L, U, Y) =>
                  L +
                  Math.max(1, String(U[X] ?? "").length) *
                    Math.max(1, String(Q[Y] ?? "").length),
                0,
              )),
              B > 2e6)
            )
              return r("#LIMIT!");
            if (D.every((L, U) => !Ce(L[X]) && al(L[X], Q[U]))) {
              if (Ce(S[X])) return S[X];
              typeof S[X] == "number" && Z.push(S[X]);
            }
          }
          return P ? (Z.length ? Zt(Z) : r("#DIV/0!")) : Nr(Z);
        }
        let R = _y(I);
        if (R && s !== "COUNT") return R;
        if (s === "TRIM")
          return String(I[0] ?? "")
            .replace(/^ +| +$/g, "")
            .replace(/ +/g, " ");
        if (s === "UPPER") {
          let T = String(I[0] ?? "").toUpperCase();
          return T.length > 32767 ? r("#LIMIT!") : T;
        }
        if (s === "VALUE") {
          if (I.length !== 1 || Array.isArray(I[0])) return r("#VALUE!");
          let T = String(I[0] ?? "").trim(),
            _ = false;
          T.startsWith("(") &&
            T.endsWith(")") &&
            ((_ = true), (T = T.slice(1, -1).trim()));
          let P = T.endsWith("%");
          if (
            (P && (T = T.slice(0, -1).trim()),
            (T = T.replace(/^\$/, "")),
            !/^[+-]?(?:\d{1,3}(?:,\d{3})+|\d+|\d*\.\d+)(?:\.\d*)?(?:[eE][+-]?\d+)?$/.test(
              T,
            ))
          )
            return r("#VALUE!");
          let S = (Number(T.replace(/,/g, "")) * (_ ? -1 : 1)) / (P ? 100 : 1);
          return Number.isFinite(S) ? S : r("#NUM!");
        }
        if (s === "MATCH") {
          if (I.length < 2 || I.length > 3) return r("#VALUE!");
          let T = I[1];
          if (
            Array.isArray(T) &&
            T.length > 1 &&
            Array.isArray(T[0]) &&
            T[0].length > 1
          )
            return r("#N/A");
          let _ = nr(T),
            P = I[0],
            S = I[2] === undefined ? 1 : Number(I[2]);
          if (![0, -1, 1].includes(S)) return r("#N/A");
          if (S === 0) {
            let K = _.findIndex((Z) =>
              typeof P == "string"
                ? typeof Z == "string" &&
                  (/[?*~]/.test(P)
                    ? al(Z, `=${P}`)
                    : Z.toLowerCase() === P.toLowerCase())
                : Z === P,
            );
            return K < 0 ? r("#N/A") : K + 1;
          }
          let F = _.map((K) => (typeof K == "string" ? K.toLowerCase() : K)),
            D = typeof P == "string" ? P.toLowerCase() : P;
          if (
            (typeof D != "number" && typeof D != "string") ||
            F.some((K) => typeof K != typeof D)
          )
            return r("#VALUE!");
          let Q = F;
          if (
            Q.some((K, Z) => Z > 0 && (S === 1 ? K < Q[Z - 1] : K > Q[Z - 1]))
          )
            return r("#VALUE!");
          let te = -1;
          return (
            Q.forEach((K, Z) => {
              (S === 1 ? K <= D : K >= D) && (te = Z);
            }),
            te < 0 ? r("#N/A") : te + 1
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
          ].includes(s)
        ) {
          let T = s === "LARGE" || s === "SMALL" ? I.slice(0, 1) : I,
            _ = [];
          T.forEach((F, D) => {
            let Q = !!c[D] && typeof c[D] == "object" && "ref" in c[D];
            for (let te of nr(F))
              typeof te == "number"
                ? _.push(te)
                : !Q &&
                  !Array.isArray(F) &&
                  (typeof te == "boolean" ||
                    (typeof te == "string" &&
                      te.trim() !== "" &&
                      Number.isFinite(Number(te)))) &&
                  _.push(Number(te));
          });
          let P = Nr(_);
          if (s === "SUM") return P;
          if (s === "AVERAGE") return _.length ? P / _.length : r("#DIV/0!");
          if (s === "MIN")
            return _.length ? _.reduce((F, D) => Math.min(F, D)) : 0;
          if (s === "MAX")
            return _.length ? _.reduce((F, D) => Math.max(F, D)) : 0;
          if (s === "STDEV.S" || s === "VAR.S") {
            if (_.length < 2) return r("#DIV/0!");
            let F =
              _.reduce((D, Q) => D + (Q - P / _.length) ** 2, 0) /
              (_.length - 1);
            return s === "VAR.S" ? F : Math.sqrt(F);
          }
          if ((_.sort((F, D) => F - D), s === "MEDIAN"))
            return _.length
              ? (_[Math.floor((_.length - 1) / 2)] +
                  _[Math.floor(_.length / 2)]) /
                  2
              : r("#NUM!");
          let S = Number(I[1]);
          return !Number.isInteger(S) || S < 1 || S > _.length
            ? r("#NUM!")
            : s === "LARGE"
              ? _[_.length - S]
              : _[S - 1];
        }
        if (s === "SUBSTITUTE") {
          if (I.length < 3 || I.length > 4 || I.some(Array.isArray))
            return r("#VALUE!");
          let T = String(I[0] ?? ""),
            _ = String(I[1] ?? ""),
            P = String(I[2] ?? ""),
            S = _.length ? T.split(_).length - 1 : 0;
          if (
            I[3] !== undefined &&
            (!Number.isInteger(Number(I[3])) || Number(I[3]) < 1)
          )
            return r("#VALUE!");
          if (
            T.length +
              (I[3] === undefined ? S : Math.min(S, 1)) *
                (P.length - _.length) >
            32767
          )
            return r("#LIMIT!");
          if (!_.length) return T;
          if (I[3] === undefined) return T.split(_).join(P);
          let F = 0;
          return T.split(_).reduce(
            (D, Q, te) =>
              te === 0 ? Q : D + (++F === Number(I[3]) ? P : _) + Q,
            "",
          );
        }
        if (s === "SEARCH" || s === "SEARCHB") {
          if (I.length < 2 || I.length > 3 || I.some(Array.isArray))
            return r("#VALUE!");
          let T = String(I[0] ?? "").toLowerCase(),
            _ = String(I[1] ?? "").toLowerCase(),
            P = I[2] === undefined ? 1 : Number(I[2]);
          if (!Number.isInteger(P) || P < 1 || P > _.length)
            return r("#VALUE!");
          if (!/[?*~]/.test(T)) {
            let S = _.indexOf(T, P - 1);
            return S < 0 ? r("#VALUE!") : S + 1;
          }
          if (_.length * _.length * T.length > 2e6) return r("#LIMIT!");
          for (let S = P - 1; S < _.length; S++)
            if (al(_.slice(S), `${T}*`)) return S + 1;
          return r("#VALUE!");
        }
        if (
          s === "CONCATENATE" &&
          I.flatMap(nr).reduce((T, _) => T + String(_ ?? "").length, 0) > 32767
        )
          return r("#LIMIT!");
        let C = l(s, c);
        return /^IM/.test(s) &&
          !["IMABS", "IMREAL", "IMAGINARY", "IMARGUMENT"].includes(s) &&
          typeof C == "number"
          ? String(C)
          : typeof C == "string" && C.length > 32767
            ? r("#LIMIT!")
            : C;
      } catch (f) {
        return Ce(f) ? r(f.error) : r("#ERROR!");
      }
    }),
    e
  );
}
function B0(e, t = {}, r, o) {
  let n = R0(e),
    i = new Set([...Pf, ...n.map(($) => $.name)]),
    a = o?.budget ?? xu(e),
    l = o?.dateContext ?? Ts(e.datePolicy),
    s = new Set(),
    c = new Map(),
    u = new Map(),
    f = new Map(),
    d = ($) => (
      [
        "#LIMIT!",
        "#BLOCKED!",
        "#CYCLE!",
        "#ERROR!",
        "#TIMEOUT!",
        "#ITER!",
      ].includes($) && s.add($),
      Oy($)
    ),
    p = kR(e, t),
    m = new Map(),
    h = new Map(),
    v = new Map(),
    g = new Map(),
    y = new Set(),
    b = new Set(),
    x = new Map(),
    k = [],
    w = [],
    A = 0,
    I = [],
    N = 0,
    R = e.definedNames?.some(($) =>
      /\b(?:_xlfn\.)?(?:LET|LAMBDA)\s*\(/i.test($.refersTo),
    ),
    C = e.definedNames?.length
      ? new RegExp(
          `(?:^|[^A-Za-z0-9_.\\\\])(?:${e.definedNames.map(($) => $.name.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")).join("|")})(?![A-Za-z0-9_.\\\\])`,
          "i",
        )
      : undefined,
    T = ($) =>
      /\b(?:_xlfn\.)?(?:LET|LAMBDA)\s*\(/i.test($) || !!(R && C?.test($)),
    _ = ER(e),
    P = [...e.sheets, ..._.sheets],
    S = new Map(P.map(($) => [$.id, $])),
    F = new Map(P.map(($) => [$.name.toLowerCase(), $])),
    D = 0,
    Q = 0,
    te = new Map(),
    K = new RegExp(
      `\\b(?:${[...ol, "ISFORMULA", "TEXTBEFORE", "TEXTAFTER", ...n.filter(($) => $.fn.result.type === "matrix").map(($) => $.name.replaceAll(".", "\\."))].join("|")})\\s*\\(`,
      "i",
    );
  for (let $ of e.sheets)
    te.set(
      $.id,
      Object.entries($.cells)
        .filter(
          ([, W]) =>
            typeof W.raw == "string" &&
            W.raw.startsWith("=") &&
            !W.literal &&
            (T(W.raw) ||
              K.test(W.raw) ||
              (/\b(?:IF|IFERROR|IFNA)\s*\(/i.test(W.raw) &&
                W.raw.replace(/"(?:[^"]|"")*"/g, '""').includes(":")) ||
              /[{}#]/.test(W.raw.replace(/"(?:[^"]|"")*"/g, '""')) ||
              /^=\s*[^()+*/]*:/.test(W.raw) ||
              (e.definedNames?.length &&
                /^=\s*[A-Za-z_][A-Za-z0-9_.]*\s*$/.test(W.raw))),
        )
        .map(([W]) => W)
        .sort((W, se) => {
          let me = je(W),
            we = je(se);
          return me.row - we.row || me.col - we.col;
        }),
    );
  let Z = ($, W, se, me = true) => {
      let we = (Ee, fe = true) => {
        let ke = xR(Ee, e, W, se, fe);
        if (ke) throw (c.set(`${W}:${se}`, ke), d(ke.error));
      };
      (($ = _.rewrite($)), we($, false));
      let ue = me ? il($, e, W, [], r?.allowName) : $;
      ((ue = _.rewrite(ue)), we(ue));
      let xe = "",
        re = 0;
      for (; re < ue.length;) {
        let Ee =
          /^(?:'((?:[^']|'')*)'|([A-Za-z_][A-Za-z0-9_.]*))!\s*(\$?[A-Za-z]{1,3}\$?[1-9]\d*)#/i.exec(
            ue.slice(re),
          );
        if (Ee) {
          let J = F.get((Ee[1]?.replace(/''/g, "'") ?? Ee[2]).toLowerCase());
          if (!J) throw d("#REF!");
          let ne = je(Ee[3]),
            ce = `${Me(ne.col)}${ne.row + 1}`,
            Te = U(J.id, ce),
            rt = v.get(`${J.id}:${ce}`);
          if (!rt) throw Ce(Te) ? Te : d("#REF!");
          ((xe += `'${J.name.replace(/'/g, "''")}'!${rt}`),
            (re += Ee[0].length));
          continue;
        }
        if (ue[re] === '"' || ue[re] === "'") {
          let J = ue[re],
            ne = re++;
          for (; re < ue.length;)
            if (ue[re++] === J)
              if (ue[re] === J) re++;
              else break;
          xe += ue.slice(ne, re);
          continue;
        }
        let fe =
          /^(SEQUENCE|FILTER|SORT|SORTBY|UNIQUE|TRANSPOSE|TAKE|DROP|TEXTSPLIT)\s*\(\s*,/i.exec(
            ue.slice(re),
          );
        if (fe) {
          ((xe += `${fe[1]}(${fe[1].toUpperCase() === "SEQUENCE" ? "1" : "#VALUE!"},`),
            (re += fe[0].length));
          continue;
        }
        let ke = /^(?:[A-Za-z_][A-Za-z0-9_.]*)?\[/.exec(ue.slice(re));
        if (ke) {
          let J = re;
          re += ke[0].length;
          let ne = 1;
          for (; re < ue.length && ne;)
            (ue[re] === "[" && ne++, ue[re] === "]" && ne--, re++);
          let ce = !ne && Lf(e, W, ue.slice(J, re), se);
          if (!ce) throw d("#REF!");
          let Te = S.get(ce.sheetId);
          xe += `'${Te.name.replace(/'/g, "''")}'!${ce.range}`;
          continue;
        }
        let z = /^(\$?[A-Za-z]{1,3}\$?[1-9]\d*)#/.exec(ue.slice(re));
        if (z) {
          let J = je(z[1]),
            ne = `${Me(J.col)}${J.row + 1}`,
            ce = U(W, ne),
            Te = v.get(`${W}:${ne}`);
          if (!Te) throw Ce(ce) ? ce : d("#REF!");
          ((xe += Te), (re += z[0].length));
          continue;
        }
        xe += ue[re++];
      }
      return xe.replace(/\b_xlfn\.(?:_xlws\.)?/gi, "");
    },
    B = ($) => {
      let W = F.get(($ ?? "").toLowerCase());
      if (!W) throw d("#REF!");
      return W;
    },
    X = () => {
      let $ = q4(
        new $y.default({
          onVariable: () => null,
          onCell: ({ sheet, row, col }) => {
            let fe = B(sheet);
            if (
              !Number.isInteger(row) ||
              !Number.isInteger(col) ||
              row < 1 ||
              col < 1 ||
              row > fe.rowCount ||
              col > fe.colCount
            )
              throw d("#REF!");
            return U(fe.id, `${Me(col - 1)}${row}`);
          },
          onRange: ({ sheet, from, to }) => {
            let fe = B(sheet),
              ke = to.row,
              z = to.col;
            if (
              from.row < 1 ||
              from.col < 1 ||
              ke > fe.rowCount ||
              z > fe.colCount ||
              ke < from.row ||
              z < from.col
            )
              throw d("#REF!");
            if ((ke - from.row + 1) * (z - from.col + 1) > 2e4)
              throw d("#LIMIT!");
            let J = a.reserve(
              (ke - from.row + 1) * (32 + (z - from.col + 1) * 8),
            );
            if (J) throw d(J);
            let ne = [];
            for (let ce = from.row; ce <= ke; ce++) {
              let Te = [];
              for (let rt = from.col; rt <= z; rt++) {
                let Le = U(fe.id, `${Me(rt - 1)}${ce}`);
                if (
                  Ce(Le) &&
                  ["#LIMIT!", "#TIMEOUT!", "#BLOCKED!"].includes(Le.error)
                )
                  throw Le;
                Te.push(Le);
              }
              ne.push(Te);
            }
            return ne;
          },
        }),
        (xe, re, Ee) =>
          p(xe, re, {
            currentSheetName: Ee.position?.sheet,
            retrieve: (fe) => Ee.retrieveRef(fe),
            issue: d,
            isIssue: Ce,
            readCell: U,
            readMetadata:
              r || o
                ? (fe, ke) => {
                    if ((o?.read(`${fe}:${ke}`), r && !r.allowRead(fe, ke)))
                      throw d("#BLOCKED!");
                    r?.step({
                      kind: "metadata",
                      sheetId: fe,
                      address: ke,
                      operation: "Inspect cell metadata",
                      result: null,
                    });
                  }
                : undefined,
            isActive: (fe, ke) => y.has(`${fe}:${ke}`),
            expandedFormula: (fe, ke) => {
              try {
                return il(String(fe.raw), e, ke);
              } catch {
                return String(fe.raw);
              }
            },
          }),
        d,
        l,
        (xe) => {
          let re = $.position;
          if (!re) return;
          let Ee = B(re.sheet).id,
            fe = `${Me(re.col - 1)}${re.row}`,
            ke = `${Ee}:${fe}:${xe.functionName}`;
          if (f.size >= 1e4 && !f.has(ke)) throw d("#LIMIT!");
          f.set(ke, { ...xe, sheetId: Ee, address: fe });
        },
        (xe, re, Ee) => {
          let fe =
            xe && typeof xe == "object" && "ref" in xe ? xe.ref : undefined;
          if (!fe?.from || !fe.to) return;
          let ke = B(fe.sheet ?? $.position?.sheet),
            z = `${Me(fe.from.col + Ee - 1)}${fe.from.row + re}`;
          if (r && !r.allowRead(ke.id, z)) throw d("#BLOCKED!");
          o?.read(`${ke.id}:${z}`);
          let J = ke.cells[z];
          if (J?.literal || typeof J?.raw != "string" || !J.raw.startsWith("="))
            return;
          let ne = J.raw;
          return (ce) => {
            if (N >= 32) throw d("#LIMIT!");
            let Te = N++;
            try {
              if (
                r?.allowDatabaseCriterion &&
                !r.allowDatabaseCriterion(ke.id, jo(ne, ce, 0))
              )
                throw d("#BLOCKED!");
              let rt = il(ne, e, ke.id, [], r?.allowName),
                Le = jo(rt, ce, 0);
              if (Le.length > en.maxFormulaLength) throw d("#LIMIT!");
              return T(Le)
                ? L(Le.slice(1), ke.id, z)
                : (I[Te] ??= X()).parse(
                    Z(Le.slice(1), ke.id, z),
                    {
                      sheet: ke.name,
                      row: fe.from.row + re,
                      col: fe.from.col + Ee,
                    },
                    true,
                  );
            } finally {
              N--;
            }
          };
        },
      );
      if (n.length) {
        let xe = $._callFunction.bind($);
        $._callFunction = (re, Ee) => {
          let fe = re.replace(/^_xlfn\./i, "").toUpperCase();
          if (!n.some((Le) => Le.name === fe)) return xe(re, Ee);
          let ke = $.position;
          if (!ke) return d("#NAME?");
          let z = B(ke.sheet).id,
            J = `${Me(ke.col - 1)}${ke.row}`,
            ne = Ee.map((Le) => $.retrieveRef(Le)),
            ce = _y(ne);
          if (ce) return ce;
          let Te = ox(e, fe, ne, z, J);
          if (Te.error) return d(Te.error);
          let rt = Te.value;
          return rt && !Array.isArray(rt) && typeof rt == "object"
            ? rt.label
            : rt;
        };
      }
      let W = (xe) => {
          let re = a.work() ?? o?.check();
          if (re) return d(re);
          let Ee = xe(),
            fe = a.value(Ee) ?? o?.check();
          return fe ? d(fe) : Ee;
        },
        se = $._callFunction.bind($),
        me = $.utils._applyInfix.bind($.utils),
        we = $.utils._applyPrefix.bind($.utils),
        ue = $.utils._applyPostfix.bind($.utils);
      if (
        (($._callFunction = (xe, re) => W(() => se(xe, re))),
        ($.utils._applyInfix = (xe, re, Ee) => W(() => me(xe, re, Ee))),
        ($.utils._applyPrefix = (xe, re, Ee) => W(() => we(xe, re, Ee))),
        ($.utils._applyPostfix = (xe, re, Ee) => W(() => ue(xe, re, Ee))),
        r)
      ) {
        let xe = (z, J, ne, ce) => {
            let Te = $.position,
              rt = B(Te.sheet).id,
              Le = `${Me(Te.col - 1)}${Te.row}`;
            try {
              let Ke = ce();
              return (
                r.step({
                  kind: z,
                  sheetId: rt,
                  address: Le,
                  operation: J,
                  inputs: ne,
                  result: Ke,
                }),
                Ke
              );
            } catch (Ke) {
              throw (
                r.step({
                  kind: z,
                  sheetId: rt,
                  address: Le,
                  operation: J,
                  inputs: ne,
                  result: Ce(Ke) ? Ke : d("#ERROR!"),
                }),
                Ke
              );
            }
          },
          re = $._callFunction.bind($),
          Ee = $.utils._applyInfix.bind($.utils),
          fe = $.utils._applyPrefix.bind($.utils),
          ke = $.utils._applyPostfix.bind($.utils);
        (($._callFunction = (z, J) => xe("function", z, J, () => re(z, J))),
          ($.utils._applyInfix = (z, J, ne) =>
            xe("operator", J, [z.val, ne.val], () => Ee(z, J, ne))),
          ($.utils._applyPrefix = (z, J, ne) =>
            xe("operator", z.join(""), [J], () => fe(z, J, ne))),
          ($.utils._applyPostfix = (z, J, ne) =>
            xe("operator", ne, [z], () => ke(z, J, ne))));
      }
      return $;
    },
    L = ($, W, se) =>
      mR($, e, W, {
        supported: i,
        issue: d,
        isIssue: Ce,
        allowName: r?.allowName,
        step: r
          ? (me, we, ue) =>
              r.step({
                kind: me,
                sheetId: W,
                address: se,
                operation: we,
                result: ue,
              })
          : undefined,
        diagnostic: (me, we) =>
          u.set(`${W}:${se}`, {
            sheetId: W,
            address: se,
            error: me,
            message: we,
          }),
        parse: (me, we, ue) => {
          let xe = Z(me, we, se, false),
            re = xe.replace(/"(?:[^"]|"")*"/g, '""');
          if (/[\[\]|]/.test(re) || /(?:https?:|file:|\\\\)/i.test(re))
            throw d("#BLOCKED!");
          if (
            Array.from(
              re
                .replace(/'(?:[^']|'')*'(?=\s*!)/g, "Sheet")
                .matchAll(/([A-Za-z_][A-Za-z0-9_.]*)\s*\(/g),
            ).some((z) => z[1] !== Uf && !i.has(z[1].toUpperCase()))
          )
            throw d("#NAME?");
          let Ee = A++,
            fe = w[Ee] ?? (w[Ee] = X()),
            ke = fe._callFunction;
          fe._callFunction = (z, J) =>
            z === Uf
              ? J.length === 1 &&
                Number.isInteger(J[0]) &&
                Number(J[0]) >= 0 &&
                Number(J[0]) < ue.length
                ? ue[Number(J[0])]
                : d("#NAME?")
              : ke(z, J);
          try {
            let z = je(se);
            return fe.parse(
              xe,
              { sheet: S.get(we).name, row: z.row + 1, col: z.col + 1 },
              true,
            );
          } finally {
            ((fe._callFunction = ke), A--);
          }
        },
      }),
    U = ($, W) => {
      let se = a.work() ?? o?.check();
      if (se) return d(se);
      if (!r) return Y($, W);
      if (!r.allowRead($, W)) return d("#BLOCKED!");
      let me = Y($, W);
      return (
        r.step({
          kind: "read",
          sheetId: $,
          address: W,
          operation: "Read cell",
          result: me,
        }),
        me
      );
    },
    Y = ($, W) => {
      if (++D > yt.maxCells) return d("#LIMIT!");
      let se, me;
      try {
        ({ row: se, col: me } = je(W));
      } catch {
        return d("#REF!");
      }
      let we = S.get($);
      if (!we || se >= we.rowCount || me >= we.colCount) return d("#REF!");
      W = `${Me(me)}${se + 1}`;
      let ue = `${$}:${W}`,
        xe = _.values.get(ue);
      if (xe) return xe.type === "error" ? d(String(xe.value)) : xe.value;
      o?.read(ue);
      let re = o?.value(ue, y.size > 0 || b.size > 0);
      if (re) return re.type === "error" ? d(String(re.value)) : re.value;
      for (let ne of y) x.get(ne)?.add(ue);
      if (y.has(ue) || b.has(ue))
        return (o?.cycle([...y, ...b, ue]), d("#CYCLE!"));
      if (m.has(ue)) return m.get(ue);
      if (h.has(ue)) return h.get(ue).value;
      let Ee = we.cells[W];
      if (!Ee || Ee.raw === null) {
        let ne = te.get($) ?? [];
        if (ne.length > 1e3) return d("#LIMIT!");
        for (let ce of ne) {
          if (ce === W) continue;
          let Te = je(ce);
          if (
            !(
              Te.row > se ||
              Te.col > me ||
              (se - Te.row + 1) * (me - Te.col + 1) > en.maxSpillCells
            ) &&
            !(y.has(`${$}:${ce}`) || b.has(`${$}:${ce}`)) &&
            (U($, ce), h.has(ue))
          )
            return h.get(ue).value;
        }
        return Ee?.raw ?? null;
      }
      if (typeof Ee.raw != "string" || !Ee.raw.startsWith("=") || Ee.literal)
        return Ee.raw;
      let fe = Ee.raw;
      o?.enter(ue);
      let z = (() => {
          if (
            y.size + b.size >= en.maxFormulaDepth ||
            fe.length > en.maxFormulaLength ||
            !Ns(fe)
          )
            return d("#LIMIT!");
          let ne,
            ce = T(fe);
          b.add(ue);
          try {
            if (((ne = ce ? fe.slice(1) : Z(fe.slice(1), $, W)), !Ns(ne)))
              return d("#LIMIT!");
          } catch (Ke) {
            return Ce(Ke)
              ? d(Ke.error)
              : d(
                  Ke instanceof Error && /^#/.test(Ke.message)
                    ? Ke.message
                    : "#NAME?",
                );
          } finally {
            b.delete(ue);
          }
          let Te = ne.replace(/"(?:[^"]|"")*"/g, '""');
          if (!ce && (/[\[\]|]/.test(Te) || /(?:https?:|file:|\\\\)/i.test(Te)))
            return d("#BLOCKED!");
          if (
            !ce &&
            Array.from(
              Te.replace(/'(?:[^']|'')*'(?=\s*!)/g, "Sheet").matchAll(
                /([A-Za-z_][A-Za-z0-9_.]*)\s*\(/g,
              ),
            ).some((Ke) => !i.has(Ke[1].toUpperCase()))
          )
            return (
              u.set(ue, {
                sheetId: $,
                address: W,
                error: "#NAME?",
                message:
                  "The formula calls an unknown or unsupported custom function. Imported VBA, add-in and external custom functions cannot execute.",
              }),
              d("#NAME?")
            );
          let rt = y.size;
          (y.add(ue), x.set(ue, new Set()));
          let Le;
          try {
            let Ke = k[rt] ?? (k[rt] = X()),
              Ct = ce
                ? L(ne, $, W)
                : Ke.parse(
                    ne,
                    { sheet: we.name, row: se + 1, col: me + 1 },
                    true,
                  );
            if (Ce(Ct)) Le = d(Ct.error);
            else if (typeof Ct == "number")
              Le = Number.isFinite(Ct) ? Ct : d("#NUM!");
            else if (Ct === null) Le = 0;
            else if (typeof Ct == "boolean") Le = Ct;
            else if (typeof Ct == "string")
              Le = Ct.length <= 32767 ? Ct : d("#LIMIT!");
            else if (Array.isArray(Ct)) {
              let Tr = Ct.every(Array.isArray) ? Ct : [Ct],
                fn = Tr[0]?.length ?? 0;
              if (!Tr.length || !fn || Tr.some((ni) => ni.length !== fn))
                Le = d("#CALC!");
              else if (
                Tr.length * fn > en.maxSpillCells ||
                Q + Tr.length * fn > yt.maxCells
              )
                Le = d("#LIMIT!");
              else if (se + Tr.length > we.rowCount || me + fn > we.colCount)
                (g.set(ue, {
                  sheetId: $,
                  anchor: W,
                  reason: "bounds",
                  message:
                    "The array extends beyond the worksheet edge. Move the formula or reduce its output.",
                }),
                  (Le = d("#SPILL!")));
              else {
                let ni = false,
                  Hy = false,
                  ll = {
                    firstRow: se,
                    lastRow: se + Tr.length - 1,
                    firstCol: me,
                    lastCol: me + fn - 1,
                  },
                  Ky = (pr) => {
                    let mt = pt(pr);
                    return (
                      mt.firstRow <= ll.lastRow &&
                      mt.lastRow >= ll.firstRow &&
                      mt.firstCol <= ll.lastCol &&
                      mt.lastCol >= ll.firstCol
                    );
                  },
                  cl = we.merges?.find(Ky),
                  Yf = we.tables?.find((pr) => Ky(pr.range));
                if (cl || Yf) {
                  ni = true;
                  let pr = cl ?? Yf.range;
                  g.set(ue, {
                    sheetId: $,
                    anchor: W,
                    reason: cl ? "merged" : "table",
                    obstruction: pr,
                    message: cl
                      ? `The array overlaps merged cells ${pr}. Unmerge those cells or move the formula.`
                      : `The array overlaps table ${Yf.name} (${pr}). Place the formula outside the table.`,
                  });
                }
                for (let pr = 0; pr < Tr.length && !ni; pr++)
                  for (let mt = 0; mt < fn; mt++) {
                    let Xr = `${Me(me + mt)}${se + pr + 1}`,
                      Ia = we.cells[Xr];
                    if (x.get(ue)?.has(`${$}:${Xr}`)) {
                      Hy = true;
                      break;
                    }
                    if (
                      Xr !== W &&
                      ((Ia && Ia.raw !== null) || h.has(`${$}:${Xr}`))
                    ) {
                      ((ni = true),
                        g.set(ue, {
                          sheetId: $,
                          anchor: W,
                          reason: "occupied",
                          obstruction: Xr,
                          message: `Cell ${Xr} blocks the array. Move the formula or review and clear the obstruction; existing evidence was preserved.`,
                        }));
                      break;
                    }
                  }
                if (Hy) Le = d("#CYCLE!");
                else if (ni) Le = d("#SPILL!");
                else {
                  let pr = (mt) =>
                    Ce(mt)
                      ? mt
                      : mt === null
                        ? 0
                        : typeof mt == "number"
                          ? Number.isFinite(mt)
                            ? mt
                            : d("#NUM!")
                          : typeof mt == "boolean"
                            ? mt
                            : typeof mt == "string"
                              ? mt.length <= 32767
                                ? mt
                                : d("#LIMIT!")
                              : d("#VALUE!");
                  for (let mt = 0; mt < Tr.length; mt++)
                    for (let Xr = 0; Xr < fn; Xr++) {
                      let Ia = `${Me(me + Xr)}${se + mt + 1}`;
                      Ia !== W &&
                        h.set(`${$}:${Ia}`, {
                          value: pr(Tr[mt][Xr]),
                          anchor: W,
                        });
                    }
                  (v.set(ue, `${W}:${Me(me + fn - 1)}${se + Tr.length}`),
                    (Q += Tr.length * fn),
                    (Le = pr(Tr[0][0])));
                }
              }
            } else Le = d("#VALUE!");
          } catch (Ke) {
            let Ct =
              Ke && typeof Ke == "object" && "details" in Ke
                ? Ke.details
                : undefined;
            Le = Ce(Ct) ? d(Ct.error) : Ce(Ke) ? d(Ke.error) : d("#ERROR!");
          } finally {
            (y.delete(ue), x.delete(ue));
          }
          return Le;
        })(),
        J = a.value(z) ?? o?.check();
      return (
        J && (z = d(J)),
        m.set(ue, z),
        o?.leave(
          ue,
          Ce(z)
            ? { type: "error", value: z.error }
            : {
                type:
                  z === null
                    ? "blank"
                    : typeof z == "string"
                      ? "text"
                      : typeof z == "boolean"
                        ? "boolean"
                        : "number",
                value: z,
              },
        ),
        r?.step({
          kind: "formula",
          sheetId: $,
          address: W,
          operation: Ee.raw,
          result: z,
        }),
        z
      );
    },
    ie = ($, W) => ie.getTypedValue($, W).value;
  ie.getTypedValue = ($, W) => {
    if (o?.running?.() === false)
      try {
        let se = je(W),
          me = S.get($);
        if (!me || se.row >= me.rowCount || se.col >= me.colCount)
          return { type: "error", value: "#REF!" };
        let we = `${Me(se.col)}${se.row + 1}`,
          ue = `${$}:${we}`,
          xe = o.value(ue, false);
        if (xe) return { ...xe };
        let re = m.get(ue) ?? h.get(ue)?.value ?? me.cells[we]?.raw ?? null;
        return Ce(re)
          ? { type: "error", value: re.error }
          : {
              type:
                re === null
                  ? "blank"
                  : typeof re == "string"
                    ? "text"
                    : typeof re == "boolean"
                      ? "boolean"
                      : "number",
              value: re,
            };
      } catch {
        return { type: "error", value: "#REF!" };
      }
    a.begin();
    try {
      D = 0;
      let se = U($, W);
      return Ce(se)
        ? { type: "error", value: se.error }
        : {
            type:
              se === null
                ? "blank"
                : typeof se == "string"
                  ? "text"
                  : typeof se == "boolean"
                    ? "boolean"
                    : "number",
            value: se,
          };
    } finally {
      a.end();
    }
  };
  let ve;
  return (
    (ie.evaluateFormula = ($, W, se) => {
      (a.begin(), (D = 0), c.delete(`${$}:${W}`), u.delete(`${$}:${W}`));
      let me = `expression:${$}:${W}`;
      try {
        let we = S.get($),
          ue = je(W);
        if (!we || ue.row >= we.rowCount || ue.col >= we.colCount)
          return "#REF!";
        if (typeof se != "string" || !se.startsWith("=")) return "#VALUE!";
        if (se.length > en.maxFormulaLength || !Ns(se))
          return d("#LIMIT!").error;
        let xe = T(se),
          re = xe ? se.slice(1) : Z(se.slice(1), $, W);
        if (!Ns(re)) return d("#LIMIT!").error;
        let Ee = re.replace(/"(?:[^"]|"")*"/g, '""');
        if (!xe && (/[\[\]|]/.test(Ee) || /(?:https?:|file:|\\\\)/i.test(Ee)))
          return d("#BLOCKED!").error;
        if (
          !xe &&
          Array.from(
            Ee.replace(/'(?:[^']|'')*'(?=\s*!)/g, "Sheet").matchAll(
              /([A-Za-z_][A-Za-z0-9_.]*)\s*\(/g,
            ),
          ).some((z) => !i.has(z[1].toUpperCase()))
        )
          return "#NAME?";
        (y.add(me), x.set(me, new Set()));
        let fe = xe
            ? L(re, $, W)
            : (ve ??= X()).parse(
                re,
                { sheet: we.name, row: ue.row + 1, col: ue.col + 1 },
                true,
              ),
          ke = a.value(fe) ?? o?.check();
        return ke
          ? d(ke).error
          : Ce(fe)
            ? d(fe.error).error
            : Array.isArray(fe)
              ? "#VALUE!"
              : fe === null
                ? 0
                : typeof fe == "number"
                  ? Number.isFinite(fe)
                    ? fe
                    : "#NUM!"
                  : typeof fe == "boolean"
                    ? fe
                    : typeof fe == "string"
                      ? fe.length <= 32767
                        ? fe
                        : d("#LIMIT!").error
                      : "#VALUE!";
      } catch (we) {
        let ue =
          we && typeof we == "object" && "details" in we
            ? we.details
            : undefined;
        return Ce(ue)
          ? d(ue.error).error
          : Ce(we)
            ? d(we.error).error
            : we instanceof Error && /^#/.test(we.message)
              ? d(we.message).error
              : d("#ERROR!").error;
      } finally {
        (a.end(), y.delete(me), x.delete(me));
      }
    }),
    (ie.getSpillRange = ($, W) => (ie($, W), v.get(`${$}:${W}`))),
    (ie.getSpillDiagnostic = ($, W) => {
      ie($, W);
      let se = g.get(`${$}:${W}`);
      return se ? { ...se } : undefined;
    }),
    (ie.getSpillAnchor = ($, W) => (ie($, W), h.get(`${$}:${W}`)?.anchor)),
    (ie.getSpillEntries = ($) => {
      a.begin();
      try {
        let W = te.get($) ?? [];
        if (W.length > 1e3)
          throw (
            d("#LIMIT!"),
            new Error("Spill discovery exceeds 1,000 possible array formulas.")
          );
        D = 0;
        for (let me of o?.running?.() === false ? [] : W)
          if ((U($, me), D > yt.maxCells))
            throw new Error(
              "Spill discovery exceeds 100,000 calculation reads.",
            );
        let se = `${$}:`;
        return [...h.entries()]
          .filter(([me]) => me.startsWith(se))
          .map(([me, { anchor }]) => ({
            address: me.slice(se.length),
            anchor,
          }));
      } finally {
        a.end();
      }
    }),
    (ie.getResourceReport = () => a.report()),
    (ie.getOperationalErrors = () => [...s]),
    (ie.getDateContext = () => ({ ...l })),
    (ie.getFinancialDiagnostics = () => structuredClone([...f.values()])),
    (ie.peekSpill = ($, W) => {
      let se = `${$}:${W}`,
        me = g.get(se);
      return { range: v.get(se), ...(me ? { diagnostic: { ...me } } : {}) };
    }),
    (ie.getNamedFormulaDiagnostics = () =>
      [...u.values()].map(($) => ({ ...$ }))),
    (ie.getReferenceDiagnostics = () => [...c.values()].map(($) => ({ ...$ }))),
    ie
  );
}
var f7 = Object.freeze({
  rows: 1e4,
  cells: 1e5,
  columns: 64,
  conditions: 128,
  values: 1e3,
  text: 1024,
  comparisonWork: 5e6,
});
var b7 = Object.freeze({ views: 20, name: 80, characters: 512e3 });
var NR = "ledgerglass-consolidation-v1",
  I7 = Object.freeze({
    sources: 8,
    categories: 3,
    contributions: 2e3,
    outputRows: 4e3,
    outputCells: 4e4,
    groups: 1e3,
    metadataCharacters: 8e6,
  }),
  Rr = E.string()
    .min(1)
    .max(128)
    .refine((e) => !!e.trim() && !/[\u0000-\u001f\u007f]/.test(e)),
  Kr = E.string().min(1).max(100),
  Uy = E.string().regex(/^[a-f0-9]{64}$/),
  RR = E.union([
    E.string().min(1).max(32767),
    E.number().finite(),
    E.boolean(),
  ]),
  Py = E.discriminatedUnion("kind", [
    E.strictObject({ kind: E.literal("constant"), value: Rr }),
    E.strictObject({ kind: E.literal("field"), field: Rr }),
  ]),
  H4 = E.strictObject({
    id: Kr,
    workbookId: Kr,
    sheetId: Kr,
    range: E.string().min(1).max(50),
    tableId: Kr.optional(),
    categories: E.array(Rr).min(1).max(3),
    amountField: Rr,
    period: Py,
    currency: Py,
    unit: Py,
    duplicatePolicy: E.enum(["reject", "sum"]),
  }),
  K4 = E.strictObject({
    engine: E.literal(NR),
    sources: E.array(H4).min(1).max(8),
    categoryLabels: E.array(Rr).min(1).max(3),
    period: Rr,
    currency: Rr,
    unit: Rr,
    linkedAmounts: E.boolean(),
    collapseDetails: E.boolean(),
  }).superRefine((e, t) => {
    let r = [
      "Row type",
      ...e.categoryLabels,
      "Period",
      "Currency",
      "Unit",
      "Amount",
      "Source",
      "Link state",
    ];
    (new Set(r.map((o) => o.toLowerCase())).size !== r.length ||
      new Set(e.sources.map((o) => o.id)).size !== e.sources.length ||
      e.sources.some(
        (o) =>
          o.categories.length !== e.categoryLabels.length ||
          new Set(o.categories).size !== o.categories.length,
      )) &&
      t.addIssue({
        code: "custom",
        message:
          "Use unique source IDs, distinct category fields, and distinct output headers with matching category counts.",
      });
  }),
  X4 = E.strictObject({
    sourceId: Kr,
    workbookId: Kr,
    workbookName: E.string().min(1).max(200),
    sheetId: Kr,
    sheetName: E.string().min(1).max(31),
    row: E.number().int().min(2).max(1048576),
    address: E.string().regex(/^[A-Z]{1,3}[1-9]\d{0,6}$/),
    locator: E.string().max(300),
    categories: E.array(RR).min(1).max(3),
    amount: E.number().finite(),
    period: Rr,
    currency: Rr,
    unit: Rr,
    link: E.enum(["live", "snapshot", "approved snapshot"]),
  }),
  J4 = E.strictObject({
    sourceId: Kr,
    row: E.number().int().positive().optional(),
    code: E.string().max(40),
    severity: E.enum(["error", "warning"]),
    message: E.string().max(4e3),
  }),
  Y4 = E.strictObject({
    sourceId: Kr,
    workbookId: Kr,
    workbookName: E.string().max(200),
    sheetId: Kr,
    sheetName: E.string().max(31),
    range: E.string().max(50),
    fingerprint: Uy,
    records: E.number().int().min(0).max(500),
    acceptedRecords: E.number().int().min(0).max(500),
    amount: E.number().finite(),
    periods: E.array(Rr).max(500),
    currencies: E.array(Rr).max(500),
    units: E.array(Rr).max(500),
  }),
  Q4 = E.strictObject({
    categories: E.array(RR).min(1).max(3),
    level: E.number().int().min(1).max(3),
    count: E.number().int().min(1).max(2e3),
    amount: E.number().finite(),
  }),
  eD = E.strictObject({
    sources: E.array(Y4).max(8),
    contributions: E.array(X4).max(2e3),
    subtotals: E.array(Q4).max(1e3),
    exceptions: E.array(J4).max(8e3),
    grandTotal: E.number().finite().nullable(),
    sourceTotal: E.number().finite().nullable(),
    difference: E.number().finite().nullable(),
    complete: E.boolean(),
  }),
  k7 = E.strictObject({
    engine: E.literal(NR),
    definition: K4,
    review: eD,
    origin: E.strictObject({
      workbookId: Kr,
      revision: E.number().int().nonnegative(),
      at: E.string().datetime(),
    }),
    evidenceHash: Uy,
    outputRange: E.string().max(50),
    outputHash: Uy,
  });
var CR = "ledgerglass-sources-v1",
  tD = ["database", "odbc", "odata", "api", "erp", "file-folder"],
  Ff = Object.freeze({
    sources: 20,
    columns: 50,
    rows: 5e3,
    cells: 1e5,
    responseBytes: 8 * 1024 * 1024,
    timeoutMs: 3e4,
    metadataCharacters: 2e5,
  }),
  zf = E.string().regex(/^[A-Za-z0-9][A-Za-z0-9._-]{0,99}$/),
  Wf = E.string()
    .min(1)
    .max(128)
    .refine((e) => e === e.trim() && !/[\u0000-\u001f\u007f]/.test(e)),
  sl = E.string().datetime({ offset: true }),
  TR = E.string().regex(/^[a-f0-9]{64}$/),
  rD = E.strictObject({
    name: Wf,
    type: E.enum(["text", "number", "boolean", "date"]),
    nullable: E.boolean(),
  }),
  Ly = E.strictObject({
    id: zf,
    label: Wf,
    kind: E.enum(tD),
    schemaVersion: zf,
    columns: E.array(rD).min(1).max(Ff.columns),
    controlFields: E.array(Wf).max(Ff.columns),
  }).superRefine((e, t) => {
    (new Set(e.columns.map((r) => r.name.toLowerCase())).size !==
      e.columns.length ||
      new Set(e.controlFields).size !== e.controlFields.length ||
      e.controlFields.some(
        (r) => !e.columns.some((o) => o.name === r && o.type === "number"),
      )) &&
      t.addIssue({
        code: "custom",
        message: "Use unique columns and declared numeric control fields.",
      });
  }),
  T7 = E.strictObject({ source: Ly, expiresAt: sl }),
  nD = E.union([
    E.string().max(32767),
    E.number().finite(),
    E.boolean(),
    E.null(),
  ]),
  C7 = E.strictObject({
    protocol: E.literal(CR),
    source: Ly,
    sourceTimestamp: sl,
    version: zf,
    rows: E.array(E.array(nD).max(Ff.columns)).max(Ff.rows),
    complete: E.literal(true),
  }),
  oD = [
    "unavailable",
    "refused",
    "expired",
    "schema",
    "limit",
    "timeout",
    "cancelled",
    "invalid",
    "stale",
  ];
var iD = E.strictObject({
    field: Wf,
    sum: E.number().finite(),
    numeric: E.number().int().min(0).max(5e3),
    blanks: E.number().int().min(0).max(5e3),
  }),
  S7 = E.strictObject({
    protocol: E.literal(CR),
    source: Ly,
    lastSuccess: E.strictObject({
      at: sl,
      sourceTimestamp: sl,
      version: zf,
      rows: E.number().int().min(0).max(5e3),
      controls: E.array(iD).max(50),
      dataHash: TR,
    }),
    lastAttempt: E.strictObject({
      at: sl,
      outcome: E.enum(["succeeded", "failed", "cancelled"]),
      code: E.enum(oD).optional(),
    }),
    outputHash: TR,
  });
var aD = "ledgerglass-analysis-v1",
  M7 = Object.freeze({
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
  Rn = E.number().finite(),
  wa = Rn.min(-1e9).max(1e9),
  Bf = E.strictObject({
    sheetId: E.string().min(1).max(100),
    address: E.string().regex(/^[A-Z]{1,3}[1-9]\d{0,6}$/),
  }),
  SR = E.string()
    .trim()
    .min(1)
    .max(31)
    .refine(
      (e) => !/[\\/?*\[\]:\u0000-\u001f]/.test(e),
      "Invalid output worksheet name",
    ),
  sD = E.strictObject({
    kind: E.literal("solver"),
    outputName: SR,
    objective: Bf,
    direction: E.enum(["maximize", "minimize", "target"]),
    target: wa.optional(),
    variables: E.array(
      E.strictObject({
        cell: Bf,
        lower: wa,
        upper: wa,
        domain: E.enum(["continuous", "integer", "binary"]),
      }),
    )
      .min(1)
      .max(6),
    constraints: E.array(
      E.strictObject({
        left: Bf,
        operator: E.enum(["<=", ">=", "="]),
        right: E.union([wa, Bf]),
      }),
    ).max(16),
    tolerance: Rn.min(1e-10).max(1e-5).default(1e-8),
    maxCandidates: E.number().int().min(1).max(5e4).default(5e4),
  }),
  lD = E.strictObject({
    kind: E.literal("statistics"),
    outputName: SR,
    sheetId: E.string().min(1).max(100),
    range: E.string().min(1).max(50),
    method: E.enum([
      "descriptive",
      "correlation",
      "covariance",
      "regression",
      "histogram",
      "distribution",
      "forecast",
    ]),
    missing: E.enum(["reject", "exclude"]).default("reject"),
    bins: E.array(wa).min(1).max(50).optional(),
    distribution: E.enum(["normal", "binomial", "poisson"]).optional(),
    mean: wa.optional(),
    standardDeviation: Rn.positive().max(1e9).optional(),
    trials: E.number().int().min(0).max(1e3).optional(),
    probability: Rn.min(0).max(1).optional(),
    lambda: Rn.positive().max(1e3).optional(),
    forecast: E.enum(["linear", "moving-average", "exponential"]).optional(),
    horizon: E.number().int().min(1).max(24).optional(),
    window: E.number().int().min(2).max(100).optional(),
    alpha: Rn.positive().max(1).optional(),
  }),
  cD = E.discriminatedUnion("kind", [sD, lD]),
  uD = E.union([E.string().max(32767), Rn, E.boolean(), E.null()]),
  fD = E.strictObject({
    status: E.enum([
      "complete",
      "optimal",
      "infeasible",
      "work-limit",
      "precision-limit",
    ]),
    success: E.boolean(),
    message: E.string().max(2e3),
    algorithm: E.string().max(300),
    assumptions: E.array(E.string().max(1e3)).max(30),
    exclusions: E.array(
      E.strictObject({
        row: E.number().int().positive(),
        reason: E.string().max(200),
        locator: E.string().max(300),
      }),
    ).max(500),
    included: E.number().int().nonnegative(),
    inspected: E.number().int().nonnegative(),
    iterations: E.number().int().nonnegative(),
    rows: E.array(E.array(uD).max(16)).max(3e3),
    solution: E.array(Rn).max(6).optional(),
    objective: Rn.optional(),
    baseline: E.array(Rn).max(6).optional(),
  }),
  Dy = E.string().regex(/^[a-f0-9]{64}$/),
  P7 = E.strictObject({
    engine: E.literal(aD),
    request: cD,
    result: fD,
    definitionHash: Dy,
    sourceHash: Dy,
    outputHash: Dy,
    outputRange: E.string().max(50),
    createdAt: E.string().datetime(),
    sourceRevision: E.number().int().nonnegative(),
    includeHidden: E.boolean(),
    applied: E.boolean(),
  });
var mD = "ledgerglass-report-output-v1",
  B7 = Object.freeze({
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
  Dr = (e) => E.string().max(e),
  Qo = E.string().min(1).max(200),
  Vf = E.string().regex(/^[a-f0-9]{64}$/),
  $R = E.string().datetime();
function pD(e) {
  if (
    !/^\$?[A-Za-z]{1,3}\$?[1-9]\d{0,6}(?::\$?[A-Za-z]{1,3}\$?[1-9]\d{0,6})?$/.test(
      e,
    )
  )
    throw Error("Choose one explicit report range.");
  let t = e.replaceAll("$", "").toUpperCase().split(":"),
    r = pt(e);
  if (
    t[0] !== `${Me(r.firstCol)}${r.firstRow + 1}` ||
    (t[1] && t[1] !== `${Me(r.lastCol)}${r.lastRow + 1}`)
  )
    throw Error("Report ranges cannot be reversed.");
  if (
    r.lastCol - r.firstCol + 1 > 20 ||
    r.lastRow - r.firstRow + 1 > 500 ||
    (r.lastCol - r.firstCol + 1) * (r.lastRow - r.firstRow + 1) > 2e3
  )
    throw Error(
      "Reports support at most 500 rows, 20 columns and 2,000 cells.",
    );
  return `${Me(r.firstCol)}${r.firstRow + 1}:${Me(r.lastCol)}${r.lastRow + 1}`;
}
var _R = E.object({
    column: E.number().int().min(0).max(16383),
    type: E.enum(["value", "text", "number", "date"]),
    operator: E.enum([
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
    value: E.union([
      E.string().max(1024),
      E.number().finite(),
      E.boolean(),
      E.null(),
    ]).optional(),
  }).strict(),
  hD = E.object({
    range: Dr(40),
    columns: E.array(
      E.object({
        column: E.number().int().min(0).max(16383),
        values: E.array(
          E.union([Dr(32767), E.number().finite(), E.boolean(), E.null()]),
        )
          .max(1e3)
          .optional(),
        conditions: E.array(_R).max(128).optional(),
        join: E.enum(["and", "or"]).optional(),
      }).strict(),
    ).max(20),
    advanced: E.array(E.array(_R).max(128)).max(128).optional(),
    caseSensitive: E.boolean().optional(),
    unique: E.boolean().optional(),
  }).strict(),
  gD = E.object({
    name: E.string().trim().min(1).max(100),
    kind: E.enum(["range", "object"]),
    sheetId: Qo,
    range: Dr(40).transform(pD),
    objectId: Qo.optional(),
    mode: E.enum(["snapshot", "linked"]),
    population: E.enum(["full", "visible"]),
    filter: hD.optional(),
  }).strict(),
  vD = E.object({
    address: E.string().regex(/^[A-Z]{1,3}[1-9]\d{0,6}$/),
    value: E.union([Dr(32767), E.number().finite(), E.boolean(), E.null()]),
    text: Dr(32767),
    style: E.unknown().optional(),
  }).strict(),
  V7 = E.object({
    engine: E.literal(mD),
    id: Qo,
    request: gD,
    workbookId: Qo,
    workbookName: Dr(200),
    revision: E.number().int().nonnegative(),
    sheetName: Dr(31),
    calculatedAt: $R,
    sourceHash: Vf,
    includeHidden: E.boolean(),
    authorizedSheetIds: E.array(Qo).max(1e3),
    sourceCells: E.array(
      E.object({
        sheetId: Qo,
        address: E.string().regex(/^[A-Z]{1,3}[1-9]\d{0,6}$/),
      }).strict(),
    ).max(4e3),
    rows: E.array(
      E.object({
        row: E.number().int().min(1).max(1048576),
        cells: E.array(vD).max(20),
      }).strict(),
    ).max(500),
    columns: E.array(E.number().int().min(0).max(16383)).max(20),
    widths: E.array(E.number().finite().min(40).max(600)).max(20),
    excludedRows: E.array(E.number().int().min(1).max(1048576)).max(500),
    object: E.unknown().optional(),
    warnings: E.array(Dr(2e3)).max(50),
    hash: Vf,
  }).strict(),
  j7 = E.object({
    id: Qo,
    fileName: Dr(255),
    fileBytes: E.number()
      .int()
      .min(1)
      .max(10 * 1024 * 1024),
    sourceSha256: Vf,
    inspectedAt: $R,
    execution: E.literal("never"),
    items: E.array(
      E.object({
        path: Dr(4096),
        kind: Dr(100),
        bytes: E.number()
          .int()
          .nonnegative()
          .max(50 * 1024 * 1024),
        sha256: Vf.optional(),
        execution: E.literal("never"),
        note: Dr(2e3),
      }).strict(),
    )
      .min(1)
      .max(1e3),
    warnings: E.array(Dr(2e3)).max(100),
  }).strict();
var Fy = Object.freeze({
    expressionCharacters: 4096,
    syntaxNodes: 512,
    depth: 48,
    calculatedColumns: 16,
  }),
  G7 = Object.freeze({
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
var yD = "ledgerglass-columnar-v1",
  bo = Object.freeze({
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
  Ea = E.string()
    .min(1)
    .max(100)
    .regex(/^[A-Za-z0-9_.-]+$/),
  ei = E.string()
    .trim()
    .min(1)
    .max(128)
    .refine((e) => !/[\u0000-\u001f\u007f]/.test(e)),
  bD = E.strictObject({
    name: ei,
    type: E.enum(["text", "number", "boolean", "date"]),
    nullable: E.boolean(),
    expression: E.string()
      .trim()
      .min(1)
      .max(Fy.expressionCharacters)
      .optional(),
  }),
  wD = E.strictObject({
    id: Ea,
    name: ei,
    sheetId: E.string().min(1).max(100),
    range: E.string().min(1).max(50),
    tableId: E.string().min(1).max(100).optional(),
    columns: E.array(bD).min(1).max(bo.columns),
    dateColumn: ei.optional(),
  }),
  jf = E.strictObject({ tableId: Ea, column: ei }),
  ED = E.strictObject({
    id: Ea,
    one: jf,
    many: jf,
    cardinality: E.enum(["one-to-many", "one-to-one"]),
    direction: E.literal("single"),
  }),
  xD = E.strictObject({
    id: Ea,
    name: ei,
    tableId: Ea,
    column: ei.optional(),
    operation: E.enum([
      "sum",
      "count",
      "countRows",
      "average",
      "min",
      "max",
      "distinctCount",
      "dax",
    ]),
    expression: E.string()
      .trim()
      .min(1)
      .max(Fy.expressionCharacters)
      .optional(),
  }),
  ID = E.strictObject({
    engine: E.literal(yD),
    name: ei,
    tables: E.array(wD).min(1).max(bo.tables),
    relationships: E.array(ED).max(bo.relationships),
    measures: E.array(xD).max(bo.measures),
  }),
  kD = E.union([
    E.string().max(32767),
    E.number().finite(),
    E.boolean(),
    E.null(),
  ]),
  AD = E.strictObject({
    ...jf.shape,
    operator: E.enum([
      "equals",
      "notEquals",
      "gt",
      "gte",
      "lt",
      "lte",
      "blank",
      "notBlank",
    ]),
    value: kD.optional(),
  }),
  ND = E.strictObject({
    groupBy: E.array(jf).max(bo.groupFields),
    filters: E.array(AD).max(bo.filters),
    measureIds: E.array(Ea).max(bo.measures),
  });
var zy = E.string().regex(/^[a-f0-9]{64}$/),
  J7 = E.strictObject({
    definition: ID,
    query: ND,
    sourceHash: zy,
    definitionHash: zy,
    outputHash: zy,
    outputRange: E.string().min(1).max(50),
    createdAt: E.string().datetime(),
    sourceRevision: E.number().int().nonnegative(),
    includeHidden: E.boolean(),
    totals: E.array(E.number().finite().nullable()).max(bo.measures),
  });
var xa = Object.freeze({
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
  OR = ["json", "xml", "pdf", "png", "jpeg", "html"],
  Yt = (e) => E.string().max(e),
  MR = E.string().regex(/^[a-f0-9]{64}$/),
  PR = E.union([Yt(32767), E.number().finite(), E.boolean(), E.null()]),
  RD = E.object({
    page: E.number().int().min(1).max(1e4),
    x: E.number().min(0).max(1),
    y: E.number().min(0).max(1),
    width: E.number().positive().max(1),
    height: E.number().positive().max(1),
  })
    .strict()
    .refine(
      (e) => e.x + e.width <= 1.000001 && e.y + e.height <= 1.000001,
      "Source region must fit its page.",
    );
function TD(e, t) {
  if (typeof e != "string") return false;
  try {
    let r = JSON.parse(e);
    return t === "array"
      ? Array.isArray(r)
      : !!r && typeof r == "object" && !Array.isArray(r);
  } catch {
    return false;
  }
}
var CD = E.object({
    value: PR,
    type: E.enum(["string", "number", "boolean", "null", "array", "object"]),
    path: Yt(2048),
    lexical: Yt(32767).optional(),
    declaredType: Yt(200).optional(),
    missing: E.boolean().optional(),
    confidence: E.number().min(0).max(1),
    region: RD.optional(),
  })
    .strict()
    .refine(
      (e) =>
        e.type === "array" || e.type === "object"
          ? TD(e.value, e.type)
          : e.type === "null"
            ? e.value === null
            : typeof e.value === e.type,
      "Native value and type disagree.",
    )
    .refine(
      (e) => !e.missing || (e.value === null && e.type === "null"),
      "Missing fields must be null placeholders.",
    ),
  UR = E.object({
    id: Yt(200).min(1),
    name: Yt(200).min(1),
    path: Yt(2048),
    ambiguous: E.boolean(),
    columns: E.array(
      E.object({ name: Yt(200).min(1), path: Yt(2048) }).strict(),
    )
      .min(1)
      .max(xa.columns),
    rows: E.array(E.array(CD).max(xa.columns)).min(1).max(xa.rows),
  })
    .strict()
    .superRefine((e, t) => {
      (new Set(e.columns.map((r) => r.path)).size !== e.columns.length ||
        e.rows.some((r) => r.length !== e.columns.length)) &&
        t.addIssue({
          code: "custom",
          message: "Table columns and rows must align with unique paths.",
        });
    }),
  Wy = E.object({
    fileName: Yt(255)
      .min(1)
      .refine((e) => !/[\\/\x00-\x1f\x7f\u202a-\u202e\u2066-\u2069]/.test(e)),
    format: E.enum(OR),
    byteLength: E.number().int().min(1).max(xa.bytes),
    sha256: MR,
    base64: Yt(Math.ceil(xa.bytes / 3) * 4),
  }).strict(),
  SD = E.object({
    schemaVersion: E.literal(1),
    fileName: Wy.shape.fileName,
    format: E.enum(OR),
    byteLength: Wy.shape.byteLength,
    sha256: MR,
    tables: E.array(UR).max(xa.tables),
    method: E.enum(["native-json", "native-xml", "document-adapter"]),
    extractor: Yt(200).min(1),
    pageCount: E.number().int().min(1).max(1e4).optional(),
    warnings: E.array(Yt(2048)).max(30),
  }).strict(),
  _D = E.object({
    decision: E.enum(["accept", "exclude", "correct"]),
    value: PR.optional(),
  })
    .strict()
    .refine(
      (e) => e.decision !== "correct" || e.value !== undefined,
      "A corrected value is required.",
    ),
  tj = E.object({
    id: Yt(100).min(1),
    sheetId: Yt(100).min(1),
    sheetName: Yt(31).min(1),
    importedAt: E.string().datetime(),
    original: Wy,
    method: SD.shape.method,
    extractor: Yt(200).min(1),
    pageCount: E.number().int().min(1).max(1e4).optional(),
    table: UR,
    resolutions: E.record(E.string().regex(/^\d+:\d+$/), _D),
    ambiguityReviewed: E.boolean(),
  }).strict();
var oj = Object.freeze({
  maxBytes: 10 * 1024 * 1024,
  maxParts: 2e3,
  maxDetails: 1e5,
});
var fj = Object.freeze({
  maxFileBytes: 10 * 1024 * 1024,
  maxSourceBytes: 256 * 1024,
  maxArtifactBytes: 2 * 1024 * 1024,
  maxStoredBytes: 10 * 1024 * 1024,
  maxArtifacts: 20,
  maxRuns: 100,
  maxCommands: 100,
  maxCells: 500,
  maxTokens: 2e4,
});
var Zf = E.number().int().min(0).max(49),
  Kn = E.array(Zf)
    .min(1)
    .max(50)
    .refine(
      (e) => new Set(e).size === e.length,
      "Choose distinct column positions.",
    ),
  LR = E.string().max(32767),
  Gf = E.union([LR, E.number().finite(), E.boolean(), E.null()]),
  DR = E.enum(["none", "upper", "lower", "proper"]),
  MD = E.strictObject({
    column: Zf,
    mode: E.enum(["whole", "token", "slice", "before", "after"]),
    delimiter: E.string().min(1).max(16).default(" "),
    index: E.number().int().min(-50).max(49).default(0),
    start: E.number().int().min(0).max(32767).default(0),
    length: E.number().int().min(1).max(32767).default(1),
  }),
  PD = E.strictObject({
    parts: E.array(MD).min(1).max(8),
    separator: E.string().max(100),
    prefix: E.string().max(1e3),
    suffix: E.string().max(1e3),
    casing: DR,
  }),
  UD = E.strictObject({ row: E.number().int().min(1).max(1048576), value: LR }),
  LD = E.array(UD)
    .min(2)
    .max(10)
    .refine(
      (e) => new Set(e.map((t) => t.row)).size === e.length,
      "Use examples from distinct source rows.",
    ),
  DD = E.discriminatedUnion("kind", [
    E.strictObject({
      kind: E.literal("trim"),
      columns: Kn,
      mode: E.enum(["edges", "collapse"]),
    }),
    E.strictObject({ kind: E.literal("clean"), columns: Kn }),
    E.strictObject({ kind: E.literal("case"), columns: Kn, mode: DR }),
    E.strictObject({
      kind: E.literal("replace"),
      columns: Kn,
      find: E.string().min(1).max(1e3),
      replacement: E.string().max(1e3),
    }),
    E.strictObject({
      kind: E.literal("split"),
      column: Zf,
      mode: E.enum(["delimited", "fixed"]),
      delimiters: E.string().max(16),
      collapse: E.boolean(),
      breaks: E.array(E.number().int().min(1).max(32766)).max(49),
    }),
    E.strictObject({
      kind: E.literal("concat"),
      columns: Kn,
      separator: E.string().max(100),
      name: E.string().min(1).max(100),
      skipBlank: E.boolean(),
    }),
    E.strictObject({
      kind: E.literal("convert"),
      columns: Kn,
      to: E.enum(["text", "number", "integer", "boolean", "date"]),
      decimal: E.enum([".", ","]),
      group: E.enum(["", ".", ",", " ", "'"]),
      dateOrder: E.enum(["MDY", "DMY"]),
    }),
    E.strictObject({
      kind: E.literal("fill"),
      columns: Kn,
      direction: E.enum(["down", "up", "value"]),
      value: Gf,
      groupColumns: E.array(Zf).max(50),
    }),
    E.strictObject({ kind: E.literal("select"), columns: Kn }),
    E.strictObject({
      kind: E.literal("pattern"),
      name: E.string().min(1).max(100),
      pattern: PD,
      examples: LD,
      inferenceColumns: Kn.max(8),
    }),
  ]),
  Lj = E.strictObject({
    sourceSheetId: E.string().min(1).max(100),
    sourceRange: E.string().min(1).max(50),
    header: E.boolean(),
    outputName: E.string().min(1).max(31),
    steps: E.array(DD).max(12),
  });
var qf = E.string().refine((e) => {
    try {
      let t = je(e);
      return `${Me(t.col)}${t.row + 1}` === e;
    } catch {
      return false;
    }
  }, "Invalid evidence address."),
  FD = E.strictObject({
    id: E.string().min(1).max(100),
    sourceId: E.string().min(1).max(100),
    sourceName: E.string().min(1).max(255),
    sourceSheet: E.string().min(1).max(31),
    originalRow: E.number().int().min(1).max(1048576),
    kind: E.enum(["file", "snapshot"]),
  }),
  zD = E.strictObject({
    address: qf,
    original: E.unknown().optional(),
    origin: FD.optional(),
  }),
  WD = E.strictObject({
    address: qf,
    row: E.number().int().min(1).max(1048576),
    field: E.string().min(1).max(150),
    value: Gf,
    type: E.enum(["text", "number", "boolean", "date", "blank"]),
    sources: E.array(qf).min(1).max(500),
    exception: E.boolean(),
  }),
  BD = E.strictObject({
    step: E.number().int().min(1).max(12),
    row: E.number().int().min(1).max(1048576),
    field: E.string().min(1).max(150),
    raw: Gf,
    sources: E.array(qf).min(1).max(500),
    reason: E.string().min(1).max(1e3),
  }),
  jj = E.strictObject({
    id: E.string().min(1).max(100),
    workbookId: E.string().min(1).max(100),
    sourceRevision: E.number().int().nonnegative(),
    sourceSheetId: E.string().min(1).max(100),
    sourceSheetName: E.string().min(1).max(31),
    outputSheetId: E.string().min(1).max(100),
    outputSheetName: E.string().min(1).max(31),
    createdAt: E.iso.datetime(),
    semantics: E.literal("reviewed-fields-v1"),
    request: E.unknown(),
    sources: E.array(zD).min(1).max(500),
    outputs: E.array(WD).min(1).max(5e3),
    exceptions: E.array(BD).max(5e3),
  });
var Hf = (e) => {
  throw new Error(e);
};
var VD = (e, t, r, o) =>
  typeof e == "number" && Number.isInteger(e) && e >= t && e <= r
    ? e
    : Hf(`${o} must be an integer between ${t} and ${r}.`);
function Me(e) {
  VD(e, 0, yt.maxCols - 1, "Column");
  let t = "";
  for (let r = e + 1; r > 0; r = Math.floor((r - 1) / 26))
    t = String.fromCharCode(65 + ((r - 1) % 26)) + t;
  return t;
}
function je(e) {
  let t = /^\$?([A-Za-z]{1,3})\$?([1-9]\d{0,6})$/.exec(e);
  if (!t) return Hf(`Invalid cell address: ${e}`);
  let r = 0;
  for (let n of t[1].toUpperCase()) r = r * 26 + n.charCodeAt(0) - 64;
  let o = Number(t[2]) - 1;
  return o >= yt.maxRows || r > yt.maxCols
    ? Hf(`Cell ${e} exceeds the supported workbook bounds.`)
    : { row: o, col: r - 1 };
}
function pt(e) {
  let t = e.trim().split(":");
  if (t.length > 2) return Hf(`Invalid range: ${e}`);
  let r = je(t[0]),
    o = je(t[1] ?? t[0]);
  return {
    firstRow: Math.min(r.row, o.row),
    lastRow: Math.max(r.row, o.row),
    firstCol: Math.min(r.col, o.col),
    lastCol: Math.max(r.col, o.col),
  };
}
var ri = new WeakMap(),
  ti = new WeakMap(),
  FR = 2e5,
  Kf = { type: "error", value: "#PENDING!" },
  zR = (e) => ({ type: "error", value: e }),
  jD = new RegExp(
    `\\b(?:${[...ol, "LET", "LAMBDA", "INDIRECT", "OFFSET", "ISFORMULA", "TEXTBEFORE", "TEXTAFTER"].join("|")})\\s*\\(`,
    "i",
  ),
  WR = { test: (e) => Qi(e).length > 0 },
  jy = (e, t) =>
    !!e.definedNames?.length ||
    [...t.values()].some((r) => {
      let o = r.formula.replace(/"(?:[^"]|"")*"/g, '""');
      return (
        jD.test(o) ||
        /[{}#]/.test(o) ||
        /^=\s*[^()+*/]*:/.test(o) ||
        (/\b(?:IF|IFERROR|IFNA)\s*\(/i.test(o) && o.includes(":"))
      );
    }),
  BR = new WeakMap();
function Xf(e) {
  if (!dn(e)) return VR(e);
  let t = BR.get(e);
  return (t || ((t = VR(e)), BR.set(e, t)), t);
}
var Zy = new WeakMap();
function ZD(e) {
  let t = dn(e.cells) ? Zy.get(e.cells)?.get(e.id) : undefined;
  if (t) return t;
  let r = [],
    o,
    n,
    i = () =>
      (o ??= Object.entries(e.cells).map(([s, c]) => [
        `${e.id}:${s}`,
        JSON.stringify([c.raw, !!c.literal]),
      ])),
    a = () => (n ??= JSON.stringify(i()));
  dn(e.cells) || a();
  for (let s in e.cells) {
    let c = e.cells[s],
      u = `${e.id}:${s}`;
    !c.literal &&
      typeof c.raw == "string" &&
      c.raw.startsWith("=") &&
      r.push([u, { sheetId: e.id, address: s, formula: c.raw }]);
  }
  let l = {
    inputs: i,
    formulas: r,
    signature: a,
    volatile: [...new Set(r.flatMap(([, s]) => Qi(s.formula)))],
  };
  if (dn(e.cells)) {
    let s = Zy.get(e.cells) ?? new Map();
    (s.set(e.id, l), Zy.set(e.cells, s));
  }
  return l;
}
function VR(e) {
  let t = e.sheets.map(ZD),
    r,
    o = () => (r ??= new Map(t.flatMap((f) => f.inputs()))),
    n = new Map(t.flatMap((f) => f.formulas)),
    i = _0(e.calculation ?? Iu),
    a = Gy(e),
    l,
    s,
    c = () =>
      (l ??= `[${JSON.stringify(e.revision)},${JSON.stringify(a)},[${t.map((f) => f.signature()).join(",")}]]`),
    u = () =>
      (s ??= `[${c()},${JSON.stringify(e.importedCalculationCaches ?? null)}]`);
  return (
    dn(e) || u(),
    {
      inputs: o,
      formulas: n,
      settings: i,
      metadata: a,
      volatile: [
        ...new Set(
          t
            .flatMap((f) => f.volatile)
            .concat((e.definedNames ?? []).flatMap((f) => Qi(f.refersTo))),
        ),
      ].sort(),
      inputSignature: c,
      signature: u,
    }
  );
}
function Jf(e, t, r = Xf(t)) {
  return !!e && (dn(t) || e.signature() === r.signature());
}
function Gy(e) {
  let t = _0(e.calculation ?? Iu);
  return JSON.stringify([
    e.id,
    t,
    Ji(e.datePolicy),
    e.definedNames,
    e.workbookLinks,
    e.extensions,
    wu(e),
    e.sheets.map((o) => [
      o.id,
      o.name,
      o.rowCount,
      o.colCount,
      o.hiddenRows ?? [],
      o.hiddenCols ?? [],
      !!o.hidden,
      !!o.veryHidden,
      o.tables ?? [],
      o.merges ?? [],
      o.outlineGroups ?? [],
    ]),
  ]);
}
function jR(e, t) {
  if (
    e.revision !== t.revision ||
    Gy(e) !== Gy(t) ||
    (e.importedCalculationCaches !== t.importedCalculationCaches &&
      JSON.stringify(e.importedCalculationCaches) !==
        JSON.stringify(t.importedCalculationCaches))
  )
    return false;
  for (let r = 0; r < e.sheets.length; r++) {
    let o = e.sheets[r].cells,
      n = t.sheets[r].cells;
    if (o === n) continue;
    let i = Object.keys(o);
    if (i.length !== Object.keys(n).length) return false;
    for (let a of i) {
      let l = o[a],
        s = n[a];
      if (!s || l.raw !== s.raw || !!l.literal != !!s.literal) return false;
    }
  }
  return true;
}
function By(e, t) {
  if (e === t || e.id !== t.id || t.revision <= e.revision) return;
  let r = ri.get(e);
  r && Jf(r, e) && ti.set(t, r);
}
function Vy(e, t) {
  if (e === t) return;
  let r = ri.get(e);
  r && Jf(r, e) && jR(e, t)
    ? ri.set(t, r)
    : ti.has(e) && jR(e, t) && ti.set(t, ti.get(e));
}
function My(e) {
  let t = Xf(e),
    r = ri.get(e);
  if (Jf(r, e, t)) return { calculator: r.calculator, report: r.report };
  let o = qR(e, t, r ? undefined : ti.get(e), false);
  return (
    qy(o.report),
    ri.set(e, o),
    ti.delete(e),
    { calculator: o.calculator, report: o.report }
  );
}
function ZR(e) {
  let t = Xf(e),
    r = ri.get(e),
    o = qR(e, t, Jf(r, e, t) ? r : undefined, true);
  return (
    qy(o.report),
    ri.set(e, o),
    ti.delete(e),
    { calculator: o.calculator, report: o.report }
  );
}
function qy(e) {
  !e ||
    typeof e != "object" ||
    Object.isFrozen(e) ||
    (Object.values(e).forEach(qy), Object.freeze(e));
}
function GD(e) {
  return Xf(e).inputSignature();
}
function GR(e, t) {
  let r = t.calculator,
    o = [];
  for (let n of e.sheets)
    try {
      for (let i of r.getSpillEntries(n.id))
        o.push({
          ...i,
          sheetId: n.id,
          result: r.getTypedValue(n.id, i.address),
        });
    } catch {
      if (!r.getOperationalErrors().includes("#LIMIT!"))
        throw Error("Could not capture spill results.");
    }
  return {
    inputSignature: GD(e),
    report: t.report,
    spills: o,
    spillMetadata: t.report.cells.map((n) => ({
      sheetId: n.sheetId,
      address: n.address,
      ...r.peekSpill(n.sheetId, n.address),
    })),
    references: r.getReferenceDiagnostics(),
    names: r.getNamedFormulaDiagnostics(),
    financial: r.getFinancialDiagnostics(),
    operationalErrors: r.getOperationalErrors(),
  };
}
function qR(e, t, r, o) {
  let n = Wo(),
    { formulas, settings } = t,
    l = Ts(e.datePolicy),
    s = xu(e);
  s.begin();
  let c = new Set(),
    u = new Map(),
    f = new Map(),
    d = new Map(),
    p = o
      ? "Explicit full recalculation"
      : r
        ? "Changed inputs and transitive dependents"
        : "First calculation of this snapshot";
  if (
    o ||
    !r ||
    t.metadata !== r.metadata ||
    jy(e, formulas) ||
    jy(e, r.formulas)
  )
    (formulas.forEach((B, X) => c.add(X)),
      r &&
        !o &&
        (p =
          t.metadata !== r.metadata
            ? "Calculation settings or structural metadata changed"
            : "Conservative recalculation for names, dynamic references or arrays"));
  else {
    let B = new Set(),
      X = t.inputs(),
      L = r.inputs();
    for (let ie of new Set([...X.keys(), ...L.keys()]))
      X.get(ie) !== L.get(ie) && B.add(ie);
    (formulas.forEach((ie, ve) => {
      (WR.test(ie.formula) ||
        !r.values.has(ve) ||
        ["#TIMEOUT!", "#LIMIT!"].includes(String(r.values.get(ve)?.value))) &&
        B.add(ve);
    }),
      r.report.cycles.flat().forEach((ie) => B.add(ie)));
    let U = new Map();
    r.dependencies.forEach((ie, ve) =>
      ie.forEach(($) => {
        let W = U.get($) ?? new Set();
        (W.add(ve), U.set($, W));
      }),
    );
    let Y = [...B];
    for (let ie = 0; ie < Y.length; ie++)
      for (let ve of U.get(Y[ie]) ?? []) B.has(ve) || (B.add(ve), Y.push(ve));
    B.forEach((ie) => {
      formulas.has(ie) && c.add(ie);
    });
  }
  formulas.forEach((B, X) => {
    !c.has(X) &&
      r?.values.has(X) &&
      (f.set(X, r.values.get(X)), u.set(X, new Set(r.dependencies.get(X))));
  });
  let h = new Map(f),
    v = [],
    g = new Set(),
    y = true,
    b,
    x = [...u.values()].reduce((B, X) => B + X.size, 0),
    k = () => (
      y && !b && (b = s.check()),
      y && !b && Wo() - n > settings.timeBudgetMs && (b = "#TIMEOUT!"),
      y ? b : undefined
    ),
    w = (B, X) => {
      let L = [];
      return B0(e, {}, undefined, {
        budget: s,
        running: () => y,
        dateContext: l,
        check: k,
        read(Y) {
          let ie = L.at(-1);
          if (!y || !ie) return;
          let ve = u.get(ie);
          ve.has(Y) || (x >= FR ? (b = "#LIMIT!") : (ve.add(Y), x++));
        },
        value(Y, ie) {
          return ie && X?.has(Y) ? X.get(Y) : B.get(Y);
        },
        enter(Y) {
          (L.push(Y),
            (x -= u.get(Y)?.size ?? 0),
            u.set(Y, new Set()),
            d.set(Y, (d.get(Y) ?? 0) + 1));
        },
        leave(Y, ie) {
          (f.set(Y, ie), L.pop());
        },
        cycle(Y) {
          let ie = Y.at(-1),
            ve = [...new Set(Y.slice(Y.indexOf(ie), -1))];
          (ve.forEach(($) => g.add($)),
            ve.length &&
              !v.some(
                ($) => $.length === ve.length && $.every((W) => ve.includes(W)),
              ) &&
              v.push(ve));
        },
      });
    },
    A = settings.mode === "manual" && !o;
  A &&
    c.forEach((B) => {
      (f.set(B, Kf), h.set(B, Kf), u.set(B, new Set(r?.dependencies.get(B))));
    });
  let I = w(h),
    N = () => {
      for (let [B, X] of formulas)
        f.set(B, h.get(B) ?? I.getTypedValue(X.sheetId, X.address));
    };
  A || N();
  let R = {
    status: settings.iterative ? "not-needed" : "off",
    passes: 0,
    maxDelta: null,
    message: settings.iterative
      ? "No circular calculation was required."
      : "Iteration is disabled. Circular references return #CYCLE!.",
  };
  if (!A && g.size && settings.iterative) {
    let B = new Set(g),
      X = new Map();
    u.forEach((U, Y) =>
      U.forEach((ie) => {
        let ve = X.get(ie) ?? new Set();
        (ve.add(Y), X.set(ie, ve));
      }),
    );
    let L = [...g];
    for (let U = 0; U < L.length; U++)
      for (let Y of X.get(L[U]) ?? []) B.has(Y) || (B.add(Y), L.push(Y));
    if (
      (f.forEach((U, Y) => {
        B.has(Y) || h.set(Y, U);
      }),
      (R.status = "failed"),
      jy(e, formulas) ||
        [...formulas.values()].some((U) => WR.test(U.formula)) ||
        g.size > 1e3)
    )
      R.message =
        "Iteration supports at most 1,000 deterministic scalar cycle cells. Names, arrays, dynamic references and volatile formulas must be removed from this calculation.";
    else {
      let U = new Map([...g].map((Y) => [Y, { type: "number", value: 0 }]));
      R.message = `No convergence within ${settings.maxIterations} simultaneous passes from zero.`;
      for (let Y = 0; Y < settings.maxIterations && !k(); Y++) {
        I = w(h, U);
        let ie = new Map(),
          ve = 0,
          $ = true;
        for (let W of g) {
          let se = formulas.get(W),
            me = I.getTypedValue(se.sheetId, se.address);
          if (
            (ie.set(W, me), me.type !== "number" || !Number.isFinite(me.value))
          ) {
            $ = false;
            break;
          }
          ve = Math.max(
            ve,
            Math.abs(Number(me.value) - Number(U.get(W).value)),
          );
        }
        if (
          ((R.passes = Y + 1),
          (R.maxDelta = Number.isFinite(ve) ? ve : null),
          !$)
        ) {
          R.message =
            "The circular calculation produced a nonnumeric value or calculation error.";
          break;
        }
        if (((U = ie), ve <= settings.maxChange)) {
          ((R.status = "converged"),
            (R.message = `Converged from zero in ${Y + 1} simultaneous passes; maximum change ${ve}.`));
          break;
        }
      }
      R.status === "converged" && !b && g.forEach((Y) => h.set(Y, U.get(Y)));
    }
    ((R.status !== "converged" || b) &&
      g.forEach((U) => h.set(U, zR(b ?? "#ITER!"))),
      (I = w(h)),
      N());
  }
  (k(),
    b &&
      (c.forEach((B) => f.set(B, zR(b))),
      R.status === "converged" && (R.status = "failed"),
      g.size &&
        settings.iterative &&
        (R.message = `Calculation stopped with ${b}; no intermediate result is current.`)),
    (y = false),
    s.end());
  let C = (B, X) => C.getTypedValue(B, X).value;
  Object.assign(C, I);
  let T = (B, X) => `${B}:${X.replace(/\$/g, "").toUpperCase()}`,
    _ = new Map();
  if (A && r)
    for (let B of e.sheets)
      for (let X of r.calculator.getSpillEntries(B.id))
        c.has(`${B.id}:${X.anchor}`) && _.set(`${B.id}:${X.address}`, X.anchor);
  C.getTypedValue = (B, X) => {
    let L = T(B, X);
    if (f.has(L)) return { ...f.get(L) };
    let U = e.sheets.find((ie) => ie.id === B)?.cells[
      X.replace(/\$/g, "").toUpperCase()
    ];
    if (
      U &&
      U.raw !== null &&
      (U.literal || typeof U.raw != "string" || !U.raw.startsWith("="))
    )
      return {
        type:
          typeof U.raw == "string"
            ? "text"
            : typeof U.raw == "boolean"
              ? "boolean"
              : "number",
        value: U.raw,
      };
    if (_.has(L)) return { ...Kf };
    let Y = I.getSpillAnchor(B, X);
    return Y && f.get(`${B}:${Y}`)?.type === "error"
      ? { ...f.get(`${B}:${Y}`) }
      : I.getTypedValue(B, X);
  };
  let P = I.getSpillEntries;
  ((C.getSpillEntries = (B) => [
    ...P(B),
    ...[..._]
      .filter(([X]) => X.startsWith(`${B}:`))
      .map(([X, L]) => ({ address: X.slice(B.length + 1), anchor: L })),
  ]),
    (C.getOperationalErrors = () => [
      ...new Set([
        ...I.getOperationalErrors(),
        ...[...f.values()]
          .filter(
            (B) =>
              B.type === "error" &&
              [
                "#PENDING!",
                "#TIMEOUT!",
                "#ITER!",
                "#LIMIT!",
                "#BLOCKED!",
                "#CYCLE!",
                "#ERROR!",
              ].includes(String(B.value)),
          )
          .map((B) => B.value),
      ]),
    ]));
  let S = (B, X) => [
      ...B.filter(
        (L) =>
          !c.has(`${L.sheetId}:${L.address}`) &&
          formulas.has(`${L.sheetId}:${L.address}`),
      ),
      ...X,
    ],
    F = r?.calculator.getReferenceDiagnostics() ?? [],
    D = r?.calculator.getNamedFormulaDiagnostics() ?? [],
    Q = r?.calculator.getFinancialDiagnostics() ?? [];
  ((C.getFinancialDiagnostics = () =>
    structuredClone(S(Q, I.getFinancialDiagnostics()))),
    (C.getReferenceDiagnostics = () => S(F, I.getReferenceDiagnostics())),
    (C.getNamedFormulaDiagnostics = () =>
      S(D, I.getNamedFormulaDiagnostics())));
  let te = new Map(
      r?.report.cells.map((B) => [`${B.sheetId}:${B.address}`, B]),
    ),
    K = new Map();
  for (let { sheetId, address, ...L } of C.getFinancialDiagnostics()) {
    let U = `${sheetId}:${address}`;
    K.set(U, [...(K.get(U) ?? []), L]);
  }
  let Z = [...formulas].map(([B, X]) => {
    let L = f.get(B) ?? Kf,
      U = r?.values.get(B);
    return {
      ...X,
      ...(L.value !== "#PENDING!" && K.has(B) ? { financial: K.get(B) } : {}),
      result: { ...L },
      status:
        L.type === "error"
          ? L.value === "#PENDING!"
            ? "pending"
            : "error"
          : "current",
      ...(L.value === "#PENDING!" && U && U.value !== "#PENDING!"
        ? { lastKnown: { revision: r.report.revision, result: { ...U } } }
        : {}),
      ...(L.value === "#PENDING!" &&
      U?.value === "#PENDING!" &&
      te.get(B)?.lastKnown
        ? { lastKnown: te.get(B).lastKnown }
        : {}),
      dependencies: [...(u.get(B) ?? [])].sort(),
      evaluations: d.get(B) ?? 0,
    };
  });
  return {
    ...t,
    calculator: C,
    values: f,
    dependencies: u,
    report: {
      resources: s.report(),
      volatile: {
        functions: [
          ...new Set(
            [...formulas.values()]
              .flatMap((B) => Qi(B.formula))
              .concat((e.definedNames ?? []).flatMap((B) => Qi(B.refersTo))),
          ),
        ].sort(),
        policy: lx,
      },
      execution: {
        mode: "cooperative",
        state:
          b === "#TIMEOUT!"
            ? "timeout"
            : b
              ? "failed"
              : A
                ? "pending"
                : "complete",
        message: b
          ? (s.report().reason ?? `Calculation stopped with ${b}.`)
          : "Synchronous calculation uses cooperative limits.",
      },
      dateContext: l,
      workbookId: e.id,
      revision: e.revision,
      run: (r?.report.run ?? 0) + 1,
      settings,
      reason: p,
      formulaCount: formulas.size,
      evaluated: Z.filter((B) => B.evaluations > 0).length,
      reused: formulas.size - c.size,
      pending: Z.filter((B) => B.status === "pending").length,
      elapsedMs: Math.max(0, Wo() - n),
      dependencyEdges: [...u.values()].reduce((B, X) => B + X.size, 0),
      cycles: v,
      iteration: R,
      cells: Z,
      limits: {
        dependencyEdges: FR,
        cooperativeTimeBudgetMs: settings.timeBudgetMs,
        hardInterrupt: false,
      },
    },
  };
}
function HR() {
  let e;
  return (t) => {
    let r = t.workbook;
    try {
      if (t.baseRevision !== undefined) {
        if (!e || e.id !== r.id || e.revision !== t.baseRevision)
          throw Error("Stale calculation delta");
        r = {
          ...r,
          sheets: r.sheets.map((a) => {
            let l = e.sheets.find((u) => u.id === a.id),
              s = t.removedCells?.[a.id] ?? [];
            if (!l) return a;
            if (!Object.keys(a.cells).length && !s.length)
              return { ...a, cells: l.cells };
            let c = { ...l.cells, ...a.cells };
            for (let u of s) delete c[u];
            return { ...a, cells: c };
          }),
        };
      }
      let { memoryBudgetBytes } = C0(r.calculation);
      if (S0(r, memoryBudgetBytes) > memoryBudgetBytes)
        throw Error(
          "#LIMIT! Calculation inputs exceed the working-data budget.",
        );
      (oi(r), nx(r, t.extensionAvailability ?? []), e && (Vy(e, r), By(e, r)));
      let n = t.force ? ZR(r) : My(r),
        i = GR(r, n);
      if (Eu(i, memoryBudgetBytes) > memoryBudgetBytes)
        throw Error(
          "#LIMIT! Calculation result payload exceeds the working-data budget.",
        );
      return ((e = r), { id: t.id, result: i });
    } catch (o) {
      return (
        (e = undefined),
        {
          id: t.id,
          error:
            o instanceof Error && o.message.startsWith("#LIMIT!")
              ? o.message
              : "Calculation worker failed; no result was published.",
        }
      );
    }
  };
}
var qD = HR(),
  KR = self;
KR.onmessage = ({ data }) => KR.postMessage(qD(data));
