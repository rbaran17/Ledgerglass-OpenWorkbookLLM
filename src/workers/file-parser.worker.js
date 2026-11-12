import {
  strFromU8 as Ao,
  strToU8 as za,
  zipSync as Rd,
  unzipSync as Nd,
  Unzip as gc,
  UnzipInflate as mc,
} from "fflate";
import * as X from "zod";
import __lib_saxes from "saxes";
import __lib_fast_formula_parser from "fast-formula-parser";
import __lib_exceljs from "exceljs";
import __lib_papaparse from "papaparse";
var T8 = Object.create;
var Ag = Object.defineProperty;
var A8 = Object.getOwnPropertyDescriptor;
var I8 = Object.getOwnPropertyNames;
var R8 = Object.getPrototypeOf,
  N8 = Object.prototype.hasOwnProperty;
var Mc = ((e) =>
  typeof require < "u"
    ? require
    : typeof Proxy < "u"
      ? new Proxy(e, { get: (t, n) => (typeof require < "u" ? require : t)[n] })
      : e)(function (e) {
  if (typeof require < "u") return require.apply(this, arguments);
  throw Error('Dynamic require of "' + e + '" is not supported');
});
var Fe = (e, t) => () => (t || e((t = { exports: {} }).exports, t), t.exports),
  Oo = (e, t) => {
    for (var n in t) Ag(e, n, { get: t[n], enumerable: true });
  },
  O8 = (e, t, n, i) => {
    if ((t && typeof t == "object") || typeof t == "function")
      for (let r of I8(t))
        !N8.call(e, r) &&
          r !== n &&
          Ag(e, r, {
            get: () => t[r],
            enumerable: !(i = A8(t, r)) || i.enumerable,
          });
    return e;
  };
var Bi = (e, t, n) => (
  (n = e != null ? T8(R8(e)) : {}),
  O8(
    t || !e || !e.__esModule
      ? Ag(n, "default", { value: e, enumerable: true })
      : n,
    e,
  )
);

var Mo = () => __lib_saxes;

var LI = () => __lib_fast_formula_parser;
var u8 = () => __lib_exceljs;
var d8 = () => __lib_papaparse;
var AN = Bi(Mo(), 1);
function rr(e) {
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
    n = [
      1779033703, 3144134277, 1013904242, 2773480762, 1359893119, 2600822924,
      528734635, 1541459225,
    ],
    i = new Uint8Array(Math.ceil((e.length + 9) / 64) * 64);
  (i.set(e), (i[e.length] = 128));
  let r = new DataView(i.buffer);
  r.setUint32(i.length - 4, e.length * 8);
  let o = new Uint32Array(64),
    a = (s, l) => (s >>> l) | (s << (32 - l));
  for (let s = 0; s < i.length; s += 64) {
    for (let g = 0; g < 16; g++) o[g] = r.getUint32(s + g * 4);
    for (let g = 16; g < 64; g++) {
      let b = o[g - 15],
        w = o[g - 2];
      o[g] =
        (o[g - 16] +
          (a(b, 7) ^ a(b, 18) ^ (b >>> 3)) +
          o[g - 7] +
          (a(w, 17) ^ a(w, 19) ^ (w >>> 10))) >>>
        0;
    }
    let [l, c, u, d, f, p, h, m] = n;
    for (let g = 0; g < 64; g++) {
      let b =
          (m +
            (a(f, 6) ^ a(f, 11) ^ a(f, 25)) +
            ((f & p) ^ (~f & h)) +
            t[g] +
            o[g]) >>>
          0,
        w =
          ((a(l, 2) ^ a(l, 13) ^ a(l, 22)) + ((l & c) ^ (l & u) ^ (c & u))) >>>
          0;
      ((m = h),
        (h = p),
        (p = f),
        (f = (d + b) >>> 0),
        (d = u),
        (u = c),
        (c = l),
        (l = (b + w) >>> 0));
    }
    [l, c, u, d, f, p, h, m].forEach((g, b) => (n[b] = (n[b] + g) >>> 0));
  }
  return n.map((s) => s.toString(16).padStart(8, "0")).join("");
}
var Hg = new WeakSet();
function Hn(e) {
  return !!e && typeof e == "object" && Hg.has(e);
}
function rl(e) {
  if (!e || typeof e != "object" || Hg.has(e)) return e;
  for (let t of Object.values(e)) rl(t);
  return (Object.freeze(e), Hg.add(e), e);
}
var h_ = new WeakMap();
function cf(e) {
  if (!Hn(e)) return true;
  let t = h_.get(e);
  if (t === undefined) {
    t = false;
    for (let n in e)
      if (e[n].hyperlink || e[n].style?.binding) {
        t = true;
        break;
      }
    h_.set(e, t);
  }
  return t;
}
var Kq = new TextEncoder();

var Ss = Object.freeze({
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
  p6 = "ledgerglass-recipes-v1",
  Kr = X.string()
    .trim()
    .min(1)
    .max(128)
    .refine((e) => !/[\u0000-\u001f\u007f]/.test(e)),
  wa = X.string().min(1).max(100),
  yw = X.union([
    X.string().max(32767),
    X.number().finite(),
    X.boolean(),
    X.null(),
  ]),
  jl = X.array(Kr)
    .max(50)
    .refine((e) => new Set(e).size === e.length, "Fields must be distinct."),
  vw = X.strictObject({
    sheetId: wa,
    range: X.string().min(1).max(50),
    tableId: wa.optional(),
  }),
  f6 = X.strictObject({
    field: Kr,
    operation: X.enum([
      "sum",
      "count",
      "average",
      "min",
      "max",
      "distinctCount",
    ]),
    name: Kr,
  }),
  gP = X.strictObject({
    field: Kr,
    operator: X.enum([
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
    value: yw.optional(),
  }),
  Ni = { id: wa },
  bP = X.discriminatedUnion("kind", [
    X.strictObject({ ...Ni, kind: X.literal("select"), fields: jl.min(1) }),
    X.strictObject({
      ...Ni,
      kind: X.literal("rename"),
      mappings: X.array(X.strictObject({ from: Kr, to: Kr }))
        .min(1)
        .max(50),
    }),
    X.strictObject({
      ...Ni,
      kind: X.literal("filter"),
      conditions: X.array(gP).min(1).max(20),
      join: X.enum(["and", "or"]),
    }),
    X.strictObject({
      ...Ni,
      kind: X.literal("types"),
      fields: jl.min(1),
      to: X.enum(["text", "number", "integer", "boolean", "date"]),
      decimal: X.enum([".", ","]),
      group: X.enum(["", ".", ",", " ", "'"]),
      dateOrder: X.enum(["MDY", "DMY"]),
    }),
    X.strictObject({
      ...Ni,
      kind: X.literal("replace"),
      fields: jl.min(1),
      mode: X.enum(["value", "text"]),
      find: yw,
      replacement: yw,
    }),
    X.strictObject({
      ...Ni,
      kind: X.literal("derive"),
      name: Kr,
      expression: X.string().min(1).max(1e3),
    }),
    X.strictObject({
      ...Ni,
      kind: X.literal("sort"),
      fields: X.array(
        X.strictObject({
          field: Kr,
          direction: X.enum(["ascending", "descending"]),
        }),
      )
        .min(1)
        .max(8),
    }),
    X.strictObject({
      ...Ni,
      kind: X.literal("unpivot"),
      fields: jl.min(1),
      attribute: Kr,
      value: Kr,
      keepNulls: X.boolean(),
    }),
    X.strictObject({
      ...Ni,
      kind: X.literal("group"),
      keys: jl,
      values: X.array(f6).min(1).max(20),
    }),
    X.strictObject({
      ...Ni,
      kind: X.literal("pivot"),
      keys: jl,
      column: Kr,
      value: Kr,
      operation: f6.shape.operation,
    }),
    X.strictObject({ ...Ni, kind: X.literal("append"), source: vw }),
    X.strictObject({
      ...Ni,
      kind: X.literal("merge"),
      source: vw,
      mode: X.enum(["inner", "left", "right", "full", "leftAnti", "rightAnti"]),
      keys: X.array(X.strictObject({ left: Kr, right: Kr }))
        .min(1)
        .max(8),
      prefix: Kr,
      normalization: X.strictObject({
        trim: X.boolean().optional(),
        collapseWhitespace: X.boolean().optional(),
        ignoreCase: X.boolean().optional(),
        unicode: X.enum(["NFC", "NFKC"]).optional(),
      }).optional(),
      fuzzy: X.strictObject({
        key: X.number().int().min(0).max(7),
        threshold: X.number().finite().min(0).max(1),
      }).optional(),
    }),
  ]),
  m6 = X.strictObject({
    engine: X.literal(p6),
    source: vw,
    steps: X.array(bP).max(Ss.steps),
  }).superRefine((e, t) => {
    (new Set(e.steps.map((n) => n.id)).size !== e.steps.length &&
      t.addIssue({ code: "custom", message: "Step IDs must be unique." }),
      new Set(
        [
          e.source,
          ...e.steps.flatMap((n) =>
            n.kind === "append" || n.kind === "merge" ? [n.source] : [],
          ),
        ].map((n) => JSON.stringify(n)),
      ).size > Ss.sources &&
        t.addIssue({
          code: "custom",
          message: "Use no more than eight recipe sources.",
        }));
  });
function yP(e) {
  return JSON.stringify(e, (t, n) =>
    n && typeof n == "object" && !Array.isArray(n)
      ? Object.fromEntries(
          Object.keys(n)
            .sort()
            .map((i) => [i, n[i]]),
        )
      : n,
  );
}
function Jr(e) {
  let t = yP(e);
  if (t.length > 16e6)
    throw Error(
      "Recipe fingerprint exceeds the 16-million-character evidence budget.",
    );
  return rr(new Uint8Array(new TextEncoder().encode(t)));
}
var Wl = X.string().regex(/^[0-9a-f]{64}$/),
  Pu = X.number()
    .int()
    .nonnegative()
    .max(Number.MAX_SAFE_INTEGER - 1),
  g6 = X.string().refine(
    (e) => Number.isFinite(Date.parse(e)) && new Date(e).toISOString() === e,
  ),
  vP = X.strictObject({
    version: X.number().int().min(1).max(Ss.versions),
    definition: m6,
    definitionHash: Wl,
    createdAt: g6,
  }),
  wP = X.array(
    X.strictObject({
      id: wa,
      name: Kr,
      versions: X.array(vP).min(1).max(Ss.versions),
    }),
  ).max(Ss.recipes),
  h6 = X.strictObject({
    name: Kr,
    types: X.array(X.enum(["blank", "number", "string", "boolean"])).max(4),
    declaredType: X.string().max(32).optional(),
  }),
  b6 = X.strictObject({
    field: Kr,
    numericCount: Pu,
    blankCount: Pu,
    sum: X.number().finite(),
  }),
  xP = X.strictObject({
    index: X.number().int().min(1).max(20),
    id: wa,
    kind: X.string().max(32),
    status: X.enum(["complete", "failed", "skipped"]),
    inputSchema: X.array(h6).max(50),
    inputRows: Pu,
    outputSchema: X.array(h6).max(50).optional(),
    outputRows: Pu.optional(),
    inputHash: Wl.optional(),
    outputHash: Wl.optional(),
    controls: X.array(b6).max(50).optional(),
    error: X.string().max(4e3).optional(),
    details: X.record(X.string(), X.unknown()).optional(),
  }),
  kP = X.strictObject({
    engine: X.literal(p6),
    recipeId: wa,
    recipeName: Kr,
    version: X.number().int().min(1).max(20),
    definition: m6,
    definitionHash: Wl,
    sourceWorkbookId: wa,
    sourceRevision: Pu,
    sourceHash: Wl,
    outputHash: Wl,
    outputRows: X.number().int().min(0).max(1e3),
    outputRange: X.string().min(1).max(50),
    at: g6,
    steps: X.array(xP).max(20),
    controls: X.array(b6).max(50),
    lineage: X.record(
      X.string().regex(/^[1-9]\d*$/),
      X.array(X.string().min(1).max(200)).max(4e3),
    ),
  });
function y6(e, t) {
  let n = e;
  if (n.transformationRecipes !== undefined) {
    if (JSON.stringify(n.transformationRecipes).length > Ss.metadataCharacters)
      throw Error("Saved recipe metadata exceeds 2,000,000 characters.");
    let r = wP.safeParse(n.transformationRecipes);
    if (!r.success)
      throw Error(`Invalid saved recipes: ${r.error.issues[0]?.message}`);
    if (
      new Set(r.data.map((o) => o.id)).size !== r.data.length ||
      new Set(r.data.map((o) => o.name.toLowerCase())).size !== r.data.length
    )
      throw Error("Saved recipe names and IDs must be unique.");
    for (let o of r.data)
      for (let [a, s] of o.versions.entries())
        if (s.version !== a + 1 || s.definitionHash !== Jr(s.definition))
          throw Error(
            "Saved recipe versions or definition hashes are inconsistent.",
          );
    t.transformationRecipes = r.data;
  }
  let i = 0;
  for (let r of t.sheets) {
    let o = n.sheets?.find((l) => l.id === r.id);
    if (o?.transformationRun === undefined) continue;
    if (
      ((i += JSON.stringify(o.transformationRun).length),
      i > Ss.runMetadataCharacters)
    )
      throw Error(
        "Recipe execution evidence exceeds 8,000,000 characters. Retain a native backup and remove obsolete output sheets.",
      );
    let a = kP.safeParse(o.transformationRun);
    if (!a.success)
      throw Error(
        `Invalid recipe execution evidence: ${a.error.issues[0]?.message}`,
      );
    let s = a.data;
    if (
      s.definitionHash !== Jr(s.definition) ||
      s.steps.length !== s.definition.steps.length ||
      s.steps.some(
        (l, c) =>
          l.status !== "complete" ||
          l.index !== c + 1 ||
          l.id !== s.definition.steps[c].id ||
          l.kind !== s.definition.steps[c].kind,
      ) ||
      Object.values(s.lineage).reduce((l, c) => l + c.length, 0) >
        Ss.lineageLinks ||
      Object.keys(s.lineage).some(
        (l) => Number(l) < 2 || Number(l) > s.outputRows + 1,
      ) ||
      Object.keys(s.lineage).length !== s.outputRows
    )
      throw Error("Recipe execution metadata is inconsistent.");
    r.transformationRun = s;
  }
}
var Du = (e) => e.replace(/^_xlfn\.(?:_xlws\.)?/i, "").toUpperCase(),
  mo = (e) => e.replace(/^_xlpm\./i, "").toLowerCase();
function _P(e, t = false) {
  return (
    (e = e.replace(/^_xlpm\./i, "")),
    /^[A-Za-z_\\][A-Za-z0-9_.\\]{0,254}$/.test(e) &&
      !/^(?:[A-Za-z]{1,3}[1-9]\d*|R\d+C\d+|R|C)$/i.test(e) &&
      !/^__LG/i.test(e) &&
      (!t || !e.includes("."))
  );
}
function Vl(e) {
  let t = e.startsWith("=") ? 1 : 0,
    n = () => {
      throw Error("Unbalanced or malformed formula syntax.");
    },
    i = (s) => {
      for (t++; t < e.length;)
        if (e[t++] === s)
          if (e[t] === s) t++;
          else return;
      n();
    },
    r = (s, l) => {
      let c = 0;
      do {
        let u = e[t];
        if (u === '"' || (u === "'" && s !== "[")) {
          i(u);
          continue;
        }
        (u === s && c++, u === l && c--, t++);
      } while (t < e.length && c);
      c && n();
    },
    o = (s) => {
      if (s > 128) throw Error("#LIMIT!");
      let l = [];
      for (; t < e.length && ![",", ")"].includes(e[t]);) {
        let c = t,
          u,
          d = e[t];
        if (d === '"') i(d);
        else if (d === "{") r("{", "}");
        else if (d === "[") r("[", "]");
        else if (d === "(") {
          t++;
          let f = o(s + 1);
          (e[t++] !== ")" && n(),
            (u = { kind: "group", body: f, start: c, end: t }));
        } else {
          let f,
            p = /^(?:'((?:[^']|'')+)'|([A-Za-z0-9_\\][A-Za-z0-9_.\\]*))!/.exec(
              e.slice(t),
            );
          p && ((f = p[1]?.replace(/''/g, "'") ?? p[2]), (t += p[0].length));
          let h =
            /^(?:#(?:REF!|NAME\?|VALUE!|DIV\/0!|N\/A|NUM!|NULL!|SPILL!|CALC!|CYCLE!|LIMIT!|BLOCKED!|ERROR!)|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}(?![\w.])|\$?[1-9]\d*:\$?[1-9]\d*(?![\w.])|\$?[A-Za-z]{1,3}\$?[1-9]\d*(?![\w.\\])|(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?)/i.exec(
              e.slice(t),
            );
          if (h) t += h[0].length;
          else {
            let m = /^[A-Za-z_\\][A-Za-z0-9_.\\]*/.exec(e.slice(t));
            m
              ? ((t += m[0].length),
                e[t] === "["
                  ? r("[", "]")
                  : (u = {
                      kind: "symbol",
                      name: m[0],
                      qualifier: f,
                      start: c,
                      end: t,
                    }))
              : ((f || d === "'" || d === "}") && n(), t++);
          }
        }
        if (
          ((u ??= { kind: "text", text: e.slice(c, t), start: c, end: t }),
          u.kind !== "text")
        )
          for (;;) {
            let f = /^\s*\(/.exec(e.slice(t));
            if (!f) break;
            t += f[0].length;
            let p = [];
            if (e[t] !== ")") for (; p.push(o(s + 1)), e[t] === ",";) t++;
            (e[t++] !== ")" && n(),
              (u = { kind: "call", callee: u, args: p, start: c, end: t }));
          }
        l.push(u);
      }
      return l;
    },
    a = o(0);
  return (t !== e.length && n(), a);
}
var ql = (e) => e.filter((t) => t.kind !== "text" || t.text.trim());
function Xl(e, t = false) {
  let n = ql(e),
    i = n[0];
  if (n.length !== 1 || i?.kind !== "symbol" || i.qualifier || !_P(i.name, t))
    throw Error(
      `Invalid ${t ? "LAMBDA parameter" : "LET variable"}: use a name, not a cell address${t ? ", without a period" : ""}.`,
    );
  return i;
}
function gp(e) {
  if (e.callee.kind !== "symbol" || e.callee.qualifier) return;
  let t = Du(e.callee.name);
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
  if (e.args.some((i) => !ql(i).length))
    throw Error(`${t} arguments cannot be empty.`);
  let n = new Set();
  for (let i = 0; i < e.args.length - 1; i += t === "LET" ? 2 : 1) {
    let r = mo(Xl(e.args[i], t === "LAMBDA").name);
    if (n.has(r)) throw Error(`Duplicate ${t} binding: ${r}.`);
    n.add(r);
  }
  return t;
}
function Fo(e, t, n = new Set()) {
  for (let i of e) {
    if (
      (i.kind === "symbol" && t(i, n, false, false),
      i.kind === "group" && Fo(i.body, t, n),
      i.kind !== "call")
    )
      continue;
    let r = gp(i);
    if (
      (i.callee.kind === "symbol"
        ? t(i.callee, n, true, false)
        : Fo([i.callee], t, n),
      !r)
    ) {
      i.args.forEach((a) => Fo(a, t, n));
      continue;
    }
    let o = new Set(n);
    for (let a = 0; a < i.args.length - 1; a += r === "LET" ? 2 : 1) {
      let s = Xl(i.args[a], r === "LAMBDA");
      (t(s, o, false, true),
        r === "LET" && Fo(i.args[a + 1], t, o),
        o.add(mo(s.name)));
    }
    Fo(i.args.at(-1), t, o);
  }
}
function v6(e) {
  try {
    let t = ql(Vl(e)),
      n = t[0];
    return t.length !== 1 || n.kind !== "call" || gp(n) !== "LAMBDA"
      ? undefined
      : n.args
          .slice(0, -1)
          .map((i) => Xl(i, true).name.replace(/^_xlpm\./i, ""));
  } catch {
    return;
  }
}
var bp = [
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
var Lu = Object.freeze({
    extensions: 8,
    functions: 40,
    calls: 100,
    inputCells: 2e3,
    outputCells: 2e3,
    metadataCharacters: 2e6,
    timeoutMs: 15e3,
    maxTimeoutMs: 6e4,
  }),
  zo = (e) =>
    X.string()
      .min(1)
      .max(e)
      .refine((t) => !/[\u0000-\u001f\u007f]/.test(t)),
  yp = X.string()
    .regex(/^[A-Za-z][A-Za-z0-9_-]{0,63}$/)
    .refine((e) => !["constructor", "prototype", "__proto__"].includes(e)),
  EP = X.enum(["number", "string", "boolean", "scalar"]),
  vp = X.strictObject({ type: EP, nullable: X.boolean().optional() }),
  x6 = X.strictObject({
    type: X.literal("matrix"),
    items: vp,
    maxRows: X.number().int().min(1).max(100),
    maxColumns: X.number().int().min(1).max(100),
  }),
  SP = X.strictObject({
    type: X.literal("entity"),
    fields: X.record(yp, vp).refine(
      (e) => Object.keys(e).length > 0 && Object.keys(e).length <= 20,
    ),
  }),
  CP = X.union([vp, x6, SP]),
  TP = X.enum([
    "read-arguments",
    "return-results",
    "linked-data",
    "remote-compute",
  ]),
  AP = X.strictObject({
    schemaVersion: X.literal(1),
    id: yp,
    label: zo(120),
    version: X.string()
      .regex(/^\d+\.\d+\.\d+(?:-[A-Za-z0-9.-]+)?$/)
      .max(64),
    namespace: X.string()
      .regex(/^[A-Z][A-Z0-9_]{1,23}$/)
      .refine(
        (e) => !["SQL", "REGISTER", "__LG", "_XLFN", "_XLUDF"].includes(e),
      ),
    runtime: X.enum(["host", "office-addin", "python", "r", "linked-data"]),
    permissions: X.array(TP).min(2).max(4),
    recalculation: X.strictObject({
      mode: X.literal("manual"),
      invalidateOn: X.literal("argument-values"),
      volatile: X.literal(false),
    }),
    functions: X.array(
      X.strictObject({
        id: X.string().regex(/^[A-Z][A-Z0-9_]{0,39}$/),
        description: zo(300),
        parameters: X.array(
          X.strictObject({ name: yp, schema: X.union([vp, x6]) }),
        ).max(16),
        result: CP,
      }),
    )
      .min(1)
      .max(20),
  });
function k6(e) {
  let t = AP.parse(e);
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
  if (new Set(t.functions.map((n) => n.id)).size !== t.functions.length)
    throw Error("Duplicate custom function names.");
  for (let n of t.functions) {
    if (bp.includes(`${t.namespace}.${n.id}`))
      throw Error("An extension cannot shadow a built-in function.");
    if (new Set(n.parameters.map((i) => i.name)).size !== n.parameters.length)
      throw Error("Duplicate extension parameter names.");
    if (n.result.type === "entity" && !t.permissions.includes("linked-data"))
      throw Error("Entity results require the linked-data permission.");
  }
  return t;
}
function ww(e, t) {
  let n = (i, r) => {
    if (i === null && r.nullable) return null;
    if (
      (typeof i == "number" &&
        Number.isFinite(i) &&
        (r.type === "number" || r.type === "scalar")) ||
      (typeof i == "string" &&
        i.length <= 32767 &&
        (r.type === "string" || r.type === "scalar")) ||
      (typeof i == "boolean" && (r.type === "boolean" || r.type === "scalar"))
    )
      return i;
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
      e.length * e[0].length > Lu.outputCells
    )
      throw Error("Extension matrix exceeds its declared shape or cell limit.");
    let i = e[0].length;
    return e.map((r) => {
      if (!Array.isArray(r) || r.length !== i)
        throw Error("Extension matrix must be rectangular.");
      return r.map((o) => n(o, t.items));
    });
  }
  if (t.type === "entity") {
    let i = X.strictObject({
      kind: X.literal("entity"),
      id: zo(200),
      label: zo(300),
      fields: X.record(yp, X.unknown()),
      source: zo(1e3),
      asOf: X.iso.datetime(),
    }).parse(e);
    if (
      Object.keys(i.fields).sort().join("|") !==
      Object.keys(t.fields).sort().join("|")
    )
      throw Error(
        "Linked entity fields must match the declared schema exactly.",
      );
    return {
      ...i,
      fields: Object.fromEntries(
        Object.entries(t.fields).map(([r, o]) => [r, n(i.fields[r], o)]),
      ),
    };
  }
  return n(e, t);
}
var w6 = new WeakSet();
function _6(e, t) {
  if (e.extensions === undefined) return;
  let n = new Set((t.definedNames ?? []).map((l) => l.name.toUpperCase())),
    i = e.extensions;
  if (
    Array.isArray(i?.installed) &&
    i.installed.some((l) =>
      l?.manifest?.functions?.some((c) =>
        n.has(`${l.manifest.namespace}.${c.id}`),
      ),
    )
  )
    throw Error("Extension functions cannot shadow defined names.");
  if (w6.has(e.extensions)) {
    t.extensions = e.extensions;
    return;
  }
  if (JSON.stringify(e.extensions).length > Lu.metadataCharacters)
    throw Error("Extension metadata exceeds its limit.");
  let r = X.strictObject({
      schemaVersion: X.literal(1),
      installed: X.array(
        X.strictObject({ manifest: X.unknown(), enabled: X.boolean() }),
      ).max(Lu.extensions),
      caches: X.array(
        X.strictObject({
          key: X.string().regex(/^[a-f0-9]{64}$/),
          sheetId: zo(100),
          address: X.string()
            .regex(/^[A-Z]{1,3}[1-9]\d*$/)
            .refine((l) => {
              let [, c, u] = /^([A-Z]+)(\d+)$/.exec(l);
              return (
                Number(u) <= 1048576 &&
                [...c].reduce((d, f) => d * 26 + f.charCodeAt(0) - 64, 0) <=
                  16384
              );
            }),
          formula: zo(8192),
          functionName: zo(80),
          arguments: X.array(X.unknown()).max(16),
          manifestHash: X.string().regex(/^[a-f0-9]{64}$/),
          scopeHash: X.string().regex(/^[a-f0-9]{64}$/),
          result: X.unknown(),
          calculatedAt: X.iso.datetime(),
          providerRevision: zo(200),
        }),
      ).max(Lu.calls),
    }).parse(e.extensions),
    o = r.installed.map((l) => ({ ...l, manifest: k6(l.manifest) }));
  if (
    new Set(o.map((l) => l.manifest.id)).size !== o.length ||
    new Set(o.map((l) => l.manifest.namespace)).size !== o.length ||
    o.reduce((l, c) => l + c.manifest.functions.length, 0) > Lu.functions
  )
    throw Error(
      "Extension identities/namespaces must be unique within the function limit.",
    );
  let a = r.caches.map((l) => {
    let c = o.find((d) =>
        l.functionName.startsWith(d.manifest.namespace + "."),
      ),
      u = c?.manifest.functions.find(
        (d) => `${c.manifest.namespace}.${d.id}` === l.functionName,
      );
    if (!u || !c || l.arguments.length !== u.parameters.length)
      throw Error("Extension cache does not match an installed function.");
    if (
      l.key !==
      E6(
        l.sheetId,
        l.address,
        l.formula,
        l.functionName,
        l.arguments,
        l.manifestHash,
      )
    )
      throw Error("Extension cache identity does not match its arguments.");
    return {
      ...l,
      arguments: l.arguments.map((d, f) => ww(d, u.parameters[f].schema)),
      result: ww(l.result, u.result),
    };
  });
  if (new Set(a.map((l) => l.key)).size !== a.length)
    throw Error("Duplicate extension cache entries.");
  let s = { schemaVersion: 1, installed: o, caches: a };
  (rl(s), w6.add(s), (t.extensions = s));
}
var IP = (e) => Jr(k6(e)),
  E6 = (e, t, n, i, r, o) => Jr([e, t, n, i, r, o]);
function kw(e) {
  return (e.extensions?.installed ?? []).flatMap((t) =>
    t.manifest.functions.map((n) => ({
      entry: t,
      fn: n,
      name: `${t.manifest.namespace}.${n.id}`,
    })),
  );
}
var xw = new WeakMap();
function wp(e) {
  return xw.get(e) ?? [];
}
function S6(e, t) {
  let n = xw.get(e);
  n && xw.set(t, n);
}
function RP(e, t, n, i) {
  return wp(e).some(
    (r) => r.id === t && r.manifestHash === n && r.sheetIds.includes(i),
  );
}
function C6(e, t, n, i, r) {
  let o = kw(e).find((d) => d.name === t);
  if (!o) return { error: "#NAME?" };
  let a = IP(o.entry.manifest);
  if (!o.entry.enabled || !RP(e, o.entry.manifest.id, a, i))
    return { error: "#NAME?" };
  if (n.length !== o.fn.parameters.length) return { error: "#VALUE!" };
  let s;
  try {
    s = n.map((d, f) => {
      let p = o.fn.parameters[f].schema;
      return ww(p.type === "matrix" && !Array.isArray(d) ? [[d]] : d, p);
    });
  } catch {
    return { error: "#VALUE!" };
  }
  let l = String(e.sheets.find((d) => d.id === i)?.cells[r]?.raw ?? ""),
    c = E6(i, r, l, t, s, a),
    u = e.extensions?.caches.find((d) => d.key === c);
  return u &&
    !wp(e).some(
      (d) => d.id === o.entry.manifest.id && d.scopeHash === u.scopeHash,
    )
    ? { error: "#NAME?" }
    : u
      ? { value: u.result }
      : { error: "#PENDING!" };
}
var _w = Object.freeze({
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
function xa() {
  return typeof performance < "u" ? performance.now() : Date.now();
}
function NP(e) {
  return {
    memoryBudgetBytes:
      (e?.memoryBudgetMiB ?? _w.defaultMemoryBudgetMiB) * 1024 * 1024,
    workBudget: e?.workBudget ?? _w.defaultWorkBudget,
  };
}
function A6(e, t) {
  let n = 0,
    i = (r, o) => {
      if (!(n > t)) {
        if (o > 80) {
          n = t + 1;
          return;
        }
        if (typeof r == "string") n += 24 + r.length * 2;
        else if (!r || typeof r != "object") n += 8;
        else if (Array.isArray(r)) {
          n += 32 + r.length * 8;
          for (let a of r) if ((i(a, o + 1), n > t)) break;
        } else {
          n += 64;
          for (let a in r) {
            if (!Object.prototype.hasOwnProperty.call(r, a)) continue;
            n += 32 + a.length * 2;
            let s = Object.getOwnPropertyDescriptor(r, a);
            if (!("value" in s)) {
              n = t + 1;
              break;
            }
            if ((i(s.value, o + 1), n > t)) break;
          }
        }
      }
    };
  return (i(e, 0), n);
}
var T6 = new WeakMap();
function OP(e, t) {
  let n = 1024,
    i = (r) => {
      n += A6(r, Math.max(0, t - n));
    };
  (i(e.definedNames),
    i(e.workbookLinks),
    i(e.extensions),
    i(e.calculation),
    i(e.datePolicy));
  for (let r of e.sheets) {
    i([
      r.id,
      r.name,
      r.tables,
      r.merges,
      r.hiddenRows,
      r.hiddenCols,
      r.outlineGroups,
    ]);
    let o = Hn(r.cells),
      a = o ? T6.get(r.cells) : undefined;
    if (a !== undefined) n += a;
    else {
      let s = n;
      for (let l in r.cells)
        if (((n += 128 + 2 * l.length), i(r.cells[l].raw), n > t)) return n;
      o && T6.set(r.cells, n - s);
    }
    if (n > t) return n;
  }
  return n;
}
function xp(e) {
  let t = NP(e.calculation),
    n = OP(e, t.memoryBudgetBytes),
    i = 0,
    r,
    o,
    a = 0,
    s = 0,
    l = 0,
    c = () => (
      !r &&
        n > t.memoryBudgetBytes &&
        ((r = "#LIMIT!"), (o = "Working-data memory budget exceeded")),
      !r &&
        i > t.workBudget &&
        ((r = "#LIMIT!"), (o = "Calculation work budget exceeded")),
      !r &&
        s + (l ? xa() - a : 0) > (e.calculation?.timeBudgetMs ?? 2500) &&
        ((r = "#TIMEOUT!"),
        (o = "Cooperative calculation time budget exceeded")),
      r
    );
  return {
    begin() {
      l++ === 0 && (a = xa());
    },
    end() {
      l > 0 && --l === 0 && (s += xa() - a);
    },
    check: c,
    work(u = 1) {
      return (
        r ||
        (i + Math.max(0, u) > t.workBudget
          ? ((r = "#LIMIT!"), (o = "Calculation work budget exceeded"), r)
          : ((i += Math.max(0, u)), c()))
      );
    },
    reserve(u) {
      return r || ((n += Math.max(0, u)), c());
    },
    value(u) {
      return r || ((n += A6(u, Math.max(0, t.memoryBudgetBytes - n))), c());
    },
    report: () => ({
      ...t,
      accountedBytes: n,
      workUnits: i,
      ...(r ? { stopped: r, reason: o } : {}),
    }),
  };
}
function Uu(e) {
  let t = e.replace(/"(?:[^"]|"")*"|'(?:[^']|'')*'/g, ""),
    n = 0;
  for (let i of t) {
    if (((i === "(" || i === "{") && n++, n > _w.maxNesting)) return false;
    (i === ")" || i === "}") && n--;
  }
  return true;
}
var ka = (e = "1900") => (e === "1904" ? 1462 : 0),
  Xi = (e = "1900") => 2958465 - ka(e),
  Hl = Object.freeze({
    calendar: "gregorian",
    dateSystem: "1900",
    textDateOrder: "MDY",
    timeZone: "UTC",
  }),
  _a = Object.freeze({
    maxSerial: 2958465,
    holidays: 1e3,
    businessDays: 366e3,
  });
function R6(e) {
  let t = (n) =>
    /\b(?:NOW|TODAY)\s*\(/i.test(n.replace(/"(?:[^"]|"")*"/g, '""'));
  return (
    e.sheets.some((n) =>
      Object.values(n.cells).some(
        (i) =>
          !i.literal &&
          typeof i.raw == "string" &&
          i.raw.startsWith("=") &&
          t(i.raw),
      ),
    ) || !!e.definedNames?.some((n) => t(n.refersTo))
  );
}
var Fu = 864e5,
  N6 = Date.UTC(1899, 11, 31);
function go(e = Hl) {
  let t = e;
  if (
    !t ||
    typeof t != "object" ||
    Array.isArray(t) ||
    Object.keys(t).some(
      (n) =>
        !["calendar", "dateSystem", "textDateOrder", "timeZone"].includes(n),
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
    let n = new Intl.DateTimeFormat("en-US", {
      timeZone: t.timeZone,
    }).resolvedOptions().timeZone;
    return {
      calendar: "gregorian",
      dateSystem: t.dateSystem,
      textDateOrder: t.textDateOrder,
      timeZone: n,
    };
  } catch {
    throw Error(
      `Unsupported time zone: ${t.timeZone}. Use an IANA time zone or UTC.`,
    );
  }
}
var kp = (e) => e === 1900 || (e % 4 === 0 && (e % 100 !== 0 || e % 400 === 0)),
  fi = (e, t) =>
    t === 2 ? (kp(e) ? 29 : 28) : [4, 6, 9, 11].includes(t) ? 30 : 31;
function Ea(e, t, n, i = "1900") {
  if (i === "1904") return (Date.UTC(e, t - 1, n) - Date.UTC(1904, 0, 1)) / Fu;
  if (e === 1900 && t === 2 && n === 29) return 60;
  let r = Date.UTC(e, t - 1, n);
  return (r - N6) / Fu + (r >= Date.UTC(1900, 2, 1) ? 1 : 0);
}
function _p(e, t = "1900") {
  return (
    Ea(e.getUTCFullYear(), e.getUTCMonth() + 1, e.getUTCDate(), t) +
    (e.getUTCHours() * 36e5 +
      e.getUTCMinutes() * 6e4 +
      e.getUTCSeconds() * 1e3 +
      e.getUTCMilliseconds()) /
      Fu
  );
}
function Ep(e, t = "1900") {
  let n = Math.floor(e);
  if (t === "1900" && n === 0) return { year: 1900, month: 1, day: 0 };
  if (t === "1900" && n === 60) return { year: 1900, month: 2, day: 29 };
  let i = new Date(
    t === "1904"
      ? Date.UTC(1904, 0, 1) + n * Fu
      : N6 + (n - (n > 60 ? 1 : 0)) * Fu,
  );
  return {
    year: i.getUTCFullYear(),
    month: i.getUTCMonth() + 1,
    day: i.getUTCDate(),
  };
}
function I6(e) {
  let t = /^(\d{1,2}):(\d{2})(?::(\d{2}(?:\.\d{1,9})?))?\s*(AM|PM)?$/i.exec(e);
  if (!t) return;
  let n = Number(t[1]),
    i = Number(t[2]),
    r = Number(t[3] ?? 0),
    o = t[4]?.toUpperCase();
  if (!(i > 59 || r >= 60 || (o ? n < 1 || n > 12 : n > 23)))
    return (
      o && (n = (n % 12) + (o === "PM" ? 12 : 0)),
      (n * 3600 + i * 60 + r) / 86400
    );
}
function Hi(e, t = Hl) {
  let n = e.trim(),
    i =
      /^(?:(\d{4})-(\d{1,2})-(\d{1,2})|(\d{1,2})\/(\d{1,2})\/(\d{4}))(?:[ T](.+))?$/.exec(
        n,
      );
  if (!i) return I6(n);
  let r = Number(i[1] ?? i[6]),
    o = Number(i[2] ?? i[t.textDateOrder === "MDY" ? 4 : 5]),
    a = Number(i[3] ?? i[t.textDateOrder === "MDY" ? 5 : 4]);
  if (
    r < (t.dateSystem === "1904" ? 1904 : 1900) ||
    r > 9999 ||
    o < 1 ||
    o > 12 ||
    a < 1 ||
    a > fi(r, o)
  )
    return;
  let s = i[7] ? I6(i[7]) : 0;
  return s === undefined ? undefined : Ea(r, o, a, t.dateSystem) + s;
}
function zu(e, t = new Date()) {
  let n = go(e),
    i = t.toISOString(),
    r = new Intl.DateTimeFormat("en-US-u-ca-gregory-nu-latn", {
      timeZone: n.timeZone,
      year: "numeric",
      month: "2-digit",
      day: "2-digit",
      hour: "2-digit",
      minute: "2-digit",
      second: "2-digit",
      hourCycle: "h23",
    }).formatToParts(t),
    o = (m) => Number(r.find((g) => g.type === m)?.value),
    a = o("year"),
    s = o("month"),
    l = o("day"),
    c = o("hour"),
    u = o("minute"),
    d = o("second"),
    f = t.getUTCMilliseconds(),
    p = Ea(a, s, l, n.dateSystem);
  if (p < 0 || p > Xi(n.dateSystem))
    throw Error(
      "Calculation clock is outside the supported 1900\u20139999 date range.",
    );
  let h = `${a}-${String(s).padStart(2, "0")}-${String(l).padStart(2, "0")}T${String(c).padStart(2, "0")}:${String(u).padStart(2, "0")}:${String(d).padStart(2, "0")}.${String(f).padStart(3, "0")}`;
  return Object.freeze({
    ...n,
    calculatedAt: i,
    localDateTime: h,
    utcOffsetMinutes: (Date.UTC(a, s - 1, l, c, u, d, f) - t.getTime()) / 6e4,
    todaySerial: p,
    nowSerial: p + (c * 3600 + u * 60 + d + f / 1e3) / 86400,
  });
}
var O6 = Object.freeze({
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
function Cs(e) {
  let t = e.replace(/"(?:[^"]|"")*"|'(?:[^']|'')*'(?=\s*!)/g, "");
  return [
    ...new Set(
      [...t.matchAll(/\b(?:_xlfn\.)?(RAND|RANDBETWEEN|NOW|TODAY)\s*\(/gi)].map(
        (n) => n[1].toUpperCase(),
      ),
    ),
  ];
}
function M6(e) {
  return (
    !!e.definedNames?.some((t) => Cs(t.refersTo).length) ||
    e.sheets.some((t) =>
      Object.values(t.cells).some(
        (n) =>
          !n.literal &&
          typeof n.raw == "string" &&
          n.raw.startsWith("=") &&
          Cs(n.raw).length,
      ),
    )
  );
}
var MP = [
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
  Ew = Symbol("not-date-function"),
  $P = {
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
function Bu(e, t, n) {
  if (!MP.includes(e)) return Ew;
  let i = n.clock.dateSystem,
    r = Xi(i),
    o = (_, S, O) => Ea(_, S, O, i),
    a = (_) => Ep(_, i),
    s = (_) => (((_ + ka(i) + 5) % 7) + 7) % 7,
    l = (_ = "#VALUE!") => {
      throw n.issue(_);
    },
    [c, u] = $P[e];
  (t.length < c || t.length > u) && l();
  let d = t.map(n.retrieve),
    f = (_) => {
      if (n.isIssue(_)) throw _;
      return Array.isArray(_) ? (_.length === 1 ? f(_[0]) : l()) : _;
    },
    p = (_) => {
      if (((_ = f(_)), _ == null)) return 0;
      if (typeof _ == "boolean") return Number(_);
      typeof _ != "number" &&
        (typeof _ != "string" ||
          !/^\s*[+-]?(?:\d+\.?\d*|\.\d+)(?:E[+-]?\d+)?\s*$/i.test(_)) &&
        l();
      let S = Number(_);
      return Number.isFinite(S) ? S : l("#NUM!");
    },
    h = (_) => (Number.isFinite(_) && _ >= 0 && _ < r + 1 ? _ : l("#NUM!")),
    m = (_) => {
      _ = f(_);
      let S =
        typeof _ == "string" &&
        !/^\s*[+-]?(?:\d+\.?\d*|\.\d+)(?:E[+-]?\d+)?\s*$/i.test(_)
          ? Hi(_, n.clock)
          : p(_);
      return S === undefined ? l() : h(S);
    },
    g = (_) => Math.floor(m(_)),
    b = (_, S, O = false) => {
      let E = a(_),
        R = E.year * 12 + E.month - 1 + S,
        I = Math.floor(R / 12),
        U = (((R % 12) + 12) % 12) + 1;
      return (
        (I < (i === "1904" ? 1904 : 1900) || I > 9999) && l("#NUM!"),
        h(o(I, U, O ? fi(I, U) : Math.min(E.day, fi(I, U))))
      );
    };
  if (e === "TODAY") return n.clock.todaySerial;
  if (e === "NOW") return n.clock.nowSerial;
  if (e === "DATE") {
    let _ = Math.trunc(p(d[0])),
      S = Math.trunc(p(d[1])),
      O = Math.trunc(p(d[2]));
    ((_ < 0 || _ >= 1e4) && l("#NUM!"), _ < 1900 && (_ += 1900));
    let E = _ * 12 + S - 1;
    _ = Math.floor(E / 12);
    let R = (((E % 12) + 12) % 12) + 1;
    return (
      (_ < 1899 || _ > 1e4 || Math.abs(O) > _a.maxSerial) && l("#NUM!"),
      h(o(_, R, 1) + O - 1)
    );
  }
  if (e === "DATEVALUE" || e === "TIMEVALUE") {
    let _ = f(d[0]);
    typeof _ != "string" && l();
    let S = Hi(_, n.clock);
    return (
      S === undefined && l(),
      e === "DATEVALUE" &&
        !/^(?:\d{4}-\d{1,2}-\d{1,2}|\d{1,2}\/\d{1,2}\/\d{4})(?:[ T]|$)/.test(
          _.trim(),
        ) &&
        l(),
      e === "DATEVALUE" ? Math.floor(S) : S - Math.floor(S)
    );
  }
  if (e === "TIME") {
    let [_, S, O] = d.map((E) => Math.trunc(p(E)));
    return (
      [_, S, O].some((E) => E < 0 || E > 32767) && l("#NUM!"),
      ((_ * 3600 + S * 60 + O) % 86400) / 86400
    );
  }
  if (["HOUR", "MINUTE", "SECOND"].includes(e)) {
    let _ = m(d[0]),
      S = Math.max(1e-9, Number.EPSILON * _ * 86400),
      O = Math.floor((_ - Math.floor(_)) * 86400 + S) % 86400;
    return e === "HOUR"
      ? Math.floor(O / 3600)
      : e === "MINUTE"
        ? Math.floor(O / 60) % 60
        : O % 60;
  }
  if (["DAY", "MONTH", "YEAR"].includes(e)) return a(g(d[0]))[e.toLowerCase()];
  if (e === "EDATE" || e === "EOMONTH")
    return b(g(d[0]), Math.trunc(p(d[1])), e === "EOMONTH");
  if (e === "DAYS") return g(d[0]) - g(d[1]);
  if (e === "WEEKDAY") {
    let _ = g(d[0]),
      S = d[1] == null ? 1 : Math.trunc(p(d[1])),
      O = s(_);
    if (S === 3) return O;
    let E =
      S === 1 || S === 17
        ? 6
        : S === 2 || S === 11
          ? 0
          : S >= 12 && S <= 16
            ? S - 11
            : l("#NUM!");
    return ((O - E + 7) % 7) + 1;
  }
  if (e === "WEEKNUM" || e === "ISOWEEKNUM") {
    let _ = g(d[0]),
      S = e === "ISOWEEKNUM" ? 21 : d[1] == null ? 1 : Math.trunc(p(d[1]));
    if (S === 21) {
      let R = _ + 3 - s(_),
        I = a(R).year;
      return Math.floor((R - o(I, 1, 1)) / 7) + 1;
    }
    let O =
        S === 1 || S === 17
          ? 6
          : S === 2 || S === 11
            ? 0
            : S >= 12 && S <= 16
              ? S - 11
              : l("#NUM!"),
      E = o(a(_).year, 1, 1);
    return Math.floor((_ - E + ((s(E) - O + 7) % 7)) / 7) + 1;
  }
  if (e === "DATEDIF") {
    let _ = g(d[0]),
      S = g(d[1]),
      O = String(f(d[2])).toUpperCase();
    _ > S && l("#NUM!");
    let E = a(_),
      R = a(S),
      I = (R.year - E.year) * 12 + R.month - E.month - (R.day < E.day ? 1 : 0);
    if (O === "D") return S - _;
    if (O === "Y") return Math.floor(I / 12);
    if (O === "M") return I;
    if (O === "YM") return I % 12;
    if (O === "MD") return S - b(_, I);
    if (O === "YD") {
      let U =
        R.year -
        (R.month < E.month || (R.month === E.month && R.day < E.day) ? 1 : 0);
      return S - o(U, E.month, Math.min(E.day, fi(U, E.month)));
    }
    l("#NUM!");
  }
  let w = (_, S, O) => {
    let E = a(_),
      R = a(S),
      I = E.day,
      U = R.day;
    return (
      O
        ? ((I = Math.min(I, 30)), (U = Math.min(U, 30)))
        : (I === fi(E.year, E.month) && (I = 30),
          U === 31 && I >= 30 && (U = 30),
          E.month === 2 &&
            E.day === fi(E.year, 2) &&
            R.month === 2 &&
            U === fi(R.year, 2) &&
            (U = 30)),
      (R.year - E.year) * 360 + (R.month - E.month) * 30 + U - I
    );
  };
  if (e === "DAYS360") {
    let _ = f(d[2] ?? false);
    return (
      typeof _ != "boolean" && typeof _ != "number" && l(),
      w(g(d[0]), g(d[1]), !!_)
    );
  }
  if (e === "YEARFRAC") {
    let _ = g(d[0]),
      S = g(d[1]),
      O = d[2] == null ? 0 : Math.trunc(p(d[2]));
    if (((O < 0 || O > 4) && l("#NUM!"), _ > S && ([_, S] = [S, _]), _ === S))
      return 0;
    if (O === 0 || O === 4) return w(_, S, O === 4) / 360;
    if (O === 2 || O === 3) return (S - _) / (O === 2 ? 360 : 365);
    let E = a(_),
      R = a(S),
      I = 365,
      U = o(E.year + 1, E.month, Math.min(E.day, fi(E.year + 1, E.month)));
    if (E.year === R.year) I = kp(E.year) ? 366 : 365;
    else if (S <= U)
      for (let P = E.year; P <= R.year; P++)
        kp(P) && o(P, 2, 29) >= _ && o(P, 2, 29) <= S && (I = 366);
    else I = (o(R.year + 1, 1, 1) - o(E.year, 1, 1)) / (R.year - E.year + 1);
    return (S - _) / I;
  }
  if (e.startsWith("WORKDAY") || e.startsWith("NETWORKDAYS")) {
    let _ = e.startsWith("NETWORK"),
      S = e.endsWith(".INTL"),
      O = g(d[0]),
      E = f(S ? (d[2] ?? 1) : 1),
      R;
    if (typeof E == "string") (/^[01]{7}$/.test(E) || l(), (R = E));
    else {
      let z = p(E),
        j;
      if (z >= 1 && z <= 7 && Number.isInteger(z))
        j = [(z + 4) % 7, (z + 5) % 7];
      else if (z >= 11 && z <= 17 && Number.isInteger(z)) j = [(z - 5) % 7];
      else return l("#NUM!");
      R = Array.from({ length: 7 }, (q, D) => (j.includes(D) ? "1" : "0")).join(
        "",
      );
    }
    let I = (z) => (Array.isArray(z) ? z.flatMap(I) : [z]),
      U = d[S ? 3 : 2] === undefined ? [] : I(d[S ? 3 : 2]);
    U.length > _a.holidays && l("#LIMIT!");
    let P = new Set();
    for (let z of U)
      z == null || z === "" || (typeof z == "boolean" && l(), P.add(g(z)));
    let T = [...R].filter((z) => z === "0").length,
      A = (z) => R[s(z)] === "0",
      L = (z, j) => {
        if (z > j) return 0;
        let q = j - z + 1,
          D = Math.floor(q / 7) * T;
        for (let W = 0; W < q % 7; W++) A(z + W) && D++;
        for (let W of P) W >= z && W <= j && A(W) && D--;
        return D;
      };
    if (_) {
      let z = g(d[1]);
      return (
        Math.abs(z - O) > _a.businessDays && l("#LIMIT!"),
        z < O ? -L(z, O) : L(O, z)
      );
    }
    let y = Math.trunc(p(d[1]));
    if ((T || l(), Math.abs(y) > _a.businessDays && l("#LIMIT!"), y === 0))
      return O;
    let v = y > 0,
      k = Math.abs(y),
      x = v ? r - O : O,
      C = (z) => (v ? L(O + 1, O + z) : L(O - z, O - 1));
    C(x) < k && l("#NUM!");
    let N = 1,
      F = x;
    for (; N < F;) {
      let z = Math.floor((N + F) / 2);
      C(z) >= k ? (F = z) : (N = z + 1);
    }
    return O + (v ? N : -N);
  }
  return l();
}
var hi = Object.freeze({
  cashFlows: 1e4,
  periods: 1e4,
  couponPeriods: 1e3,
  evaluations: 2048,
  minRate: -0.999999999,
  maxRate: 1e6,
  relativeResidual: 1e-11,
});
function $6(e, t, n = hi.minRate) {
  let i = Math.log1p(n),
    r = Math.log1p(hi.maxRate),
    o = Math.log1p(t),
    a = 0,
    s = [],
    l = (g) => {
      if (g < i || g > r || !Number.isFinite(g) || a >= hi.evaluations) return;
      a++;
      let b = e(Math.expm1(g));
      if (!(
        !Number.isFinite(b.value) ||
        !Number.isFinite(b.scale) ||
        b.scale <= 0
      ))
        return { ...b, normalized: b.value / b.scale };
    },
    c = (g, b) => {
      if (!b || Math.abs(b.normalized) > hi.relativeResidual) return false;
      let w = Math.expm1(g);
      return (
        s.some(
          (_) => Math.abs(_.rate - w) <= 1e-7 * Math.max(1, Math.abs(w)),
        ) ||
          s.push({
            rate: w,
            residual: b.value,
            relativeResidual: Math.abs(b.normalized),
          }),
        true
      );
    };
  if ([0, Math.log1p(0.1), Math.log1p(0.5)].map(l).every((g) => g?.value === 0))
    return {
      status: "indeterminate",
      candidates: s,
      evaluations: a,
      exhaustive: false,
    };
  let d = o;
  for (let g = 0; g < 60; g++) {
    let b = l(d);
    if (!b || c(d, b)) break;
    let w = 1e-5,
      _ = l(Math.max(i, d - w)),
      S = l(Math.min(r, d + w));
    if (!_ || !S) break;
    let O =
      (S.normalized - _.normalized) / (Math.min(r, d + w) - Math.max(i, d - w));
    if (!Number.isFinite(O) || Math.abs(O) < 1e-15) break;
    let E = d - Math.max(-1, Math.min(1, b.normalized / O));
    if (E < i || E > r || E === d) break;
    d = E;
  }
  let f = [
      ...new Set([
        ...Array.from({ length: 193 }, (g, b) => i + ((r - i) * b) / 192),
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
          .filter((g) => g >= n)
          .map(Math.log1p),
      ]),
    ].sort((g, b) => g - b),
    p;
  for (let g of f) {
    let b = l(g);
    if (!b) {
      p = undefined;
      continue;
    }
    if ((c(g, b), p && p.y.normalized * b.normalized < 0)) {
      let w = p.x,
        _ = g,
        S = p.y.normalized;
      for (let O = 0; O < 100; O++) {
        let E = (w + _) / 2,
          R = l(E);
        if (!R || c(E, R) || E === w || E === _) break;
        S * R.normalized < 0 ? (_ = E) : ((w = E), (S = R.normalized));
      }
    }
    p = { x: g, y: b };
  }
  s.sort((g, b) => g.rate - b.rate);
  let h = [...s].sort(
      (g, b) =>
        Math.abs(Math.log1p(g.rate) - o) - Math.abs(Math.log1p(b.rate) - o),
    )[0],
    m = a >= hi.evaluations;
  return {
    ...(h && !m ? { selected: h } : {}),
    candidates: s,
    evaluations: a,
    status: m ? "limit" : h ? "converged" : "no-convergence",
    exhaustive: false,
  };
}
var Cw = [
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
  Tw = Symbol("not-financial"),
  PP = {
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
  Sw = (e) => (Array.isArray(e) ? e.flatMap(Sw) : [e]),
  Sa = (e) => {
    let t = 0,
      n = 0;
    for (let i of e) {
      let r = i - n,
        o = t + r;
      ((n = o - t - r), (t = o));
    }
    return t;
  },
  DP = (e) => {
    let t = e.filter((n) => n !== 0).map(Math.sign);
    return t.slice(1).filter((n, i) => n !== t[i]).length;
  };
function P6(e, t, n) {
  if (!Cw.includes(e)) return Tw;
  let i =
      "Cash paid out is negative; cash received is positive. Rates and periods must use the same interval. No automatic cent rounding.",
    r = false,
    o = n.clock.dateSystem,
    a = (E, R, I) => Ea(E, R, I, o),
    s = (E) => Ep(E, o),
    l = (E) => {
      ((r = true), n.report?.({ functionName: e, convention: i, ...E }));
    },
    c = (E, R) => {
      throw (l({ status: "error", error: E, message: R }), n.issue(E));
    },
    u = (E) =>
      Number.isFinite(E)
        ? E
        : c(
            "#NUM!",
            "The financial equation overflowed or has no finite result. Review rates, signs and periods.",
          ),
    d = (E) => {
      if (n.isIssue(E)) throw E;
      return Array.isArray(E)
        ? E.length === 1
          ? d(E[0])
          : c("#VALUE!", "A scalar argument must contain exactly one value.")
        : E;
    },
    f = (E, R) => (
      (E = d(E)),
      E == null
        ? (R ?? 0)
        : typeof E == "boolean"
          ? Number(E)
          : typeof E != "number" &&
              (typeof E != "string" ||
                !/^\s*[+-]?(?:\d+\.?\d*|\.\d+)(?:E[+-]?\d+)?\s*$/i.test(E))
            ? c(
                "#VALUE!",
                "Expected a finite number; nonnumeric text cannot be used as a financial input.",
              )
            : u(Number(E))
    ),
    p = (E) => {
      if (((E = d(E)), E == null || typeof E == "boolean"))
        return c(
          "#VALUE!",
          "A cash-flow or security date is missing or is not a date.",
        );
      let R = typeof E == "string" ? Hi(E, n.clock) : f(E);
      return R === undefined || R < (o === "1904" ? 0 : 1) || R >= Xi(o) + 1
        ? c(
            "#VALUE!",
            "Invalid cash-flow or security date. Use a valid Gregorian/workbook-system serial or explicit full-year date text.",
          )
        : Math.floor(R);
    },
    h = (E) => (
      E <= 0 && c("#NUM!", "Periods must be greater than zero."),
      E > hi.periods &&
        c("#LIMIT!", "Financial periods exceed the 10000-period budget."),
      E
    ),
    m = (E) =>
      E > -1 ? E : c("#NUM!", "The rate must be greater than -100%."),
    g = (E) =>
      [0, 1].includes(E)
        ? E
        : c("#NUM!", "Payment timing must be 0 (end) or 1 (beginning)."),
    b = (E, R) => Math.exp(R * Math.log1p(E)),
    w = (E, R) => (E === 0 ? R : Math.expm1(R * Math.log1p(E)) / E),
    _ = (E, R, I, U, P) => -(I * b(E, R) + U) / ((1 + E * P) * w(E, R)),
    S = (E, R = false) => {
      let I = Sw(E);
      I.length > hi.cashFlows &&
        c("#LIMIT!", "Cash-flow input exceeds 10000 positions.");
      let U = [];
      for (let P of I) {
        if (n.isIssue(P)) throw P;
        typeof P == "number"
          ? U.push(u(P))
          : R &&
            c(
              "#VALUE!",
              "Dated cash flows require one numeric amount for every date; blanks, booleans and text cannot be silently removed.",
            );
      }
      return (
        U.length || c("#NUM!", "No numeric cash flows were supplied."),
        U
      );
    },
    O = (E, R, I, U) => {
      (R < (U ?? hi.minRate) || R > hi.maxRate) &&
        c("#NUM!", "The rate guess is outside the supported search interval.");
      let P = $6(E, R, U),
        T = P.status === "limit" ? "#LIMIT!" : "#NUM!";
      if (
        (l({
          status: P.selected ? "converged" : "error",
          ...(P.selected ? {} : { error: T }),
          guess: R,
          signChanges: I,
          solver: P,
          message: P.selected
            ? `Residual-checked rate. ${P.candidates.length} candidate root(s) found; selected nearest the guess in log-rate space. ${I !== undefined && I > 1 ? "Cash flows change sign more than once; multiple returns may exist. " : ""}The bounded search does not prove uniqueness.`
            : P.status === "indeterminate"
              ? "The equation does not identify a unique rate; no rate has been invented."
              : "No residual-checked rate converged within the supported search interval and evaluation budget. Review signs, timing and guess.",
        }),
        !P.selected)
      )
        throw n.issue(T);
      return P.selected.rate;
    };
  try {
    let [E, R] = PP[e];
    (t.length < E || t.length > R) &&
      c("#VALUE!", `${e} requires ${E}\u2013${R} arguments.`);
    let I = t.map(n.retrieve),
      U;
    if (["PV", "FV", "PMT", "IPMT", "PPMT", "RATE", "NPER"].includes(e))
      if (e === "RATE") {
        let P = h(f(I[0])),
          T = f(I[1]),
          A = f(I[2]),
          L = f(I[3]),
          y = g(f(I[4]));
        ([A, T, L].every((v) => v === 0) &&
          c("#NUM!", "All amounts are zero; no rate can be determined."),
          (U = O(
            (v) => {
              let k = [A * b(v, P), T * (1 + v * y) * w(v, P), L];
              return { value: Sa(k), scale: Sa(k.map(Math.abs)) };
            },
            f(I[5], 0.1),
          )));
      } else if (e === "NPER") {
        let P = m(f(I[0])),
          T = f(I[1]),
          A = f(I[2]),
          L = f(I[3]),
          y = g(f(I[4]));
        if (P === 0)
          (T === 0 &&
            c(
              "#DIV/0!",
              "Zero rate and zero payment cannot determine a period count.",
            ),
            (U = -(A + L) / T));
        else {
          let v = T * (1 + P * y),
            k = (v - L * P) / (A * P + v);
          (k > 0 ||
            c(
              "#NUM!",
              "Cash-flow signs or payment size do not produce a finite real period count.",
            ),
            (U = Math.log(k) / Math.log1p(P)));
        }
        h(U);
      } else if (e === "IPMT" || e === "PPMT") {
        let P = m(f(I[0])),
          T = f(I[1]),
          A = h(f(I[2])),
          L = f(I[3]),
          y = f(I[4]),
          v = g(f(I[5]));
        (T < 1 || T > A || !Number.isInteger(T)) &&
          c(
            "#NUM!",
            "Interest/principal period must be an integer from 1 through nper.",
          );
        let k = _(P, A, L, y, v),
          x =
            v && T === 1
              ? 0
              : (-(L * b(P, T - 1) + k * (1 + P * v) * w(P, T - 1)) * P) /
                (1 + P * v);
        U = e === "IPMT" ? x : k - x;
      } else {
        let P = m(f(I[0])),
          T = h(f(I[1])),
          A = f(I[2]),
          L = f(I[3]),
          y = g(f(I[4]));
        U =
          e === "PMT"
            ? _(P, T, A, L, y)
            : e === "FV"
              ? -(L * b(P, T) + A * (1 + P * y) * w(P, T))
              : -(L + A * (1 + P * y) * w(P, T)) / b(P, T);
      }
    else if (["NPV", "XNPV", "IRR", "XIRR", "MIRR"].includes(e)) {
      let P = e.startsWith("X"),
        T = e === "IRR" || e === "XIRR",
        A = S(e === "NPV" ? I.slice(1) : I[T || e === "MIRR" ? 0 : 1], P),
        L = A.map((v, k) => k + (e === "NPV" ? 1 : 0));
      if (P) {
        let v = Sw(I[T ? 1 : 2]);
        v.length !== A.length &&
          c("#NUM!", "Every cash-flow amount must have exactly one date.");
        let k = v.map(p);
        (k.some((x) => x < k[0]) &&
          c("#NUM!", "A cash-flow date precedes the first cash-flow date."),
          (L = k.map((x) => (x - k[0]) / 365)),
          (i =
            "Dated cash flows use truncated Gregorian/workbook-system dates and actual elapsed days / 365, including leap years. The first date is time zero; later dates may be unsorted or repeated. One numeric amount per date."));
      } else
        i =
          "Periodic cash-flow positions are equally spaced. NPV starts at period 1; IRR/MIRR start at period 0. Numeric zeros retain a period; blanks, text and booleans in cash-flow lists are ignored. No automatic cent rounding.";
      let y = (v) => {
        let k = A.map((x, C) => x / Math.exp(L[C] * Math.log1p(v)));
        return { value: Sa(k), scale: Sa(k.map(Math.abs)) };
      };
      if (T || e === "MIRR")
        if (
          ((!A.some((v) => v < 0) || !A.some((v) => v > 0)) &&
            c(
              "#NUM!",
              "At least one positive and one negative cash flow are required.",
            ),
          e === "MIRR")
        ) {
          let v = m(f(I[1])),
            k = m(f(I[2])),
            x = A.length - 1,
            C = Sa(
              A.map((F, z) => (F < 0 ? F / Math.exp(z * Math.log1p(v)) : 0)),
            ),
            N = Sa(
              A.map((F, z) =>
                F > 0 ? F * Math.exp((x - z) * Math.log1p(k)) : 0,
              ),
            );
          U = Math.expm1(Math.log(-N / C) / x);
        } else U = O(y, f(I[P ? 2 : 1], 0.1), DP(A));
      else U = y(m(f(I[0]))).value;
    } else if (["SLN", "SYD", "DB", "DDB"].includes(e)) {
      let P = f(I[0]),
        T = f(I[1]),
        A = h(f(I[2]));
      if (
        ((i =
          "Depreciation uses matching life/period units and no automatic cent rounding. DB rounds its fixed rate to three decimals and prorates first/last years; DDB does not switch to straight line."),
        (P < 0 || T < 0) && c("#NUM!", "Cost and salvage must be nonnegative."),
        e === "SLN")
      )
        U = (P - T) / A;
      else {
        let L = f(I[3]);
        if (
          ((L < 1 || L > A + (e === "DB" ? 1 : 0) || !Number.isInteger(L)) &&
            c(
              "#NUM!",
              "Depreciation period must be a supported whole period within asset life (plus DB partial final year).",
            ),
          e === "SYD")
        )
          U = ((P - T) * (A - L + 1) * 2) / (A * (A + 1));
        else if (e === "DB") {
          let y = f(I[4], 12);
          if (
            ((!Number.isInteger(A) ||
              y < 1 ||
              y > 12 ||
              !Number.isInteger(y) ||
              L > A + (y < 12 ? 1 : 0)) &&
              c(
                "#NUM!",
                "DB requires whole life/period and 1\u201312 first-year months; a partial first year permits one partial final year.",
              ),
            P === 0 || T >= P)
          )
            U = 0;
          else {
            let v = Math.round((1 - (T / P) ** (1 / A)) * 1e3) / 1e3,
              k = P;
            U = 0;
            for (let x = 1; x <= L; x++)
              ((U = k * v * (x === 1 ? y / 12 : x > A ? (12 - y) / 12 : 1)),
                (k -= U));
          }
        } else {
          let y = f(I[4], 2);
          y <= 0 && c("#NUM!", "DDB factor must be greater than zero.");
          let v = P;
          U = 0;
          for (let k = 1; k <= L; k++)
            ((U = Math.max(0, Math.min((v * y) / A, v - T))), (v -= U));
        }
      }
    } else {
      let P = e === "ACCRINT" ? 6 : e === "ACCRINTM" ? 4 : 6,
        T = Math.trunc(f(I[P]));
      ([0, 1, 2, 3, 4].includes(T) ||
        c(
          "#NUM!",
          "Day-count basis must be 0 (US 30/360), 1 (actual/actual), 2 (actual/360), 3 (actual/365), or 4 (European 30/360).",
        ),
        (i = `Security basis ${T}; annual coupon/yield rates, regular maturity-anchored coupons at frequency 1/2/4, end-of-month anchoring. PRICE/YIELD are per 100 face value. Gregorian/workbook-system dates; no business-day adjustment or automatic rounding.`));
      let A = (k, x) =>
          T === 0 || T === 4
            ? Bu("DAYS360", [k, x, T === 4], { ...n, retrieve: (C) => C })
            : x - k,
        L = (k, x, C) => (T === 1 ? x - k : (T === 3 ? 365 : 360) / C),
        y = (k) => {
          let x = Math.trunc(f(k));
          return [1, 2, 4].includes(x)
            ? x
            : c("#NUM!", "Coupon frequency must be 1, 2 or 4.");
        },
        v = (k, x) => {
          let C = s(k),
            N = C.year * 12 + C.month - 1 + x,
            F = Math.floor(N / 12),
            z = (((N % 12) + 12) % 12) + 1;
          return a(
            F,
            z,
            C.day === fi(C.year, C.month)
              ? fi(F, z)
              : Math.min(C.day, fi(F, z)),
          );
        };
      if (e === "PRICE" || e === "YIELD") {
        let k = p(I[0]),
          x = p(I[1]),
          C = f(I[2]),
          N = f(I[3]),
          F = f(I[4]),
          z = y(I[5]);
        (k >= x || C < 0 || F <= 0 || N < 0 || (e === "YIELD" && N === 0)) &&
          c(
            "#NUM!",
            "Require settlement before maturity, nonnegative coupon/yield, and positive redemption/price.",
          );
        let j = 1,
          q = v(x, -12 / z);
        for (; q > k;)
          (j >= hi.couponPeriods &&
            c("#LIMIT!", "Security exceeds 1000 coupon periods."),
            j++,
            (q = v(x, (-j * 12) / z)));
        let D = v(x, (-(j - 1) * 12) / z),
          W = L(q, D, z),
          $ = A(q, k),
          M = T === 0 || T === 4 ? W - $ : A(k, D),
          B = M / W,
          V = (100 * C) / z,
          G = (H) =>
            j === 1
              ? (F + V) / (1 + (H / z) * B) - (V * $) / W
              : Sa(
                  Array.from(
                    { length: j },
                    (K, te) => V / (1 + H / z) ** (te + B),
                  ),
                ) +
                F / (1 + H / z) ** (j - 1 + B) -
                (V * $) / W;
        U =
          e === "PRICE"
            ? G(N)
            : O(
                (H) => {
                  let K = G(H);
                  return { value: K - N, scale: Math.abs(K) + N };
                },
                0.1,
                undefined,
                0,
              );
      } else if (e === "ACCRINTM") {
        let k = p(I[0]),
          x = p(I[1]),
          C = f(I[2]),
          N = f(I[3], 1e3);
        (k >= x || C <= 0 || N <= 0) &&
          c(
            "#NUM!",
            "Require issue before settlement, positive coupon rate and positive par.",
          );
        let F = Bu("YEARFRAC", [k, x, T], { ...n, retrieve: (z) => z });
        ((U = N * C * F),
          (i = `Accrued interest at maturity: par \xD7 annual rate \xD7 YEARFRAC basis ${T}, using the published local day-count convention. Gregorian/workbook-system; no automatic rounding.`));
      } else {
        let k = p(I[0]),
          x = p(I[1]),
          C = p(I[2]),
          N = f(I[3]),
          F = f(I[4], 1e3),
          z = y(I[5]),
          j = g(f(I[7], 1));
        (k >= C || x <= k || N <= 0 || F <= 0) &&
          c(
            "#NUM!",
            "Require issue before first interest and settlement, positive rate and positive par.",
          );
        let q = !j && C > x ? x : k,
          D = Math.floor(
            (s(q).year * 12 + s(q).month - s(x).year * 12 - s(x).month) /
              (12 / z),
          );
        for (; v(x, (D * 12) / z) > q;) D--;
        let W = v(x, (D * 12) / z),
          $ = v(x, ((D + 1) * 12) / z),
          M = 0,
          B = 0;
        for (; W < C;) {
          ++B > hi.couponPeriods &&
            c("#LIMIT!", "Accrual exceeds 1000 quasi-coupon periods.");
          let V = Math.max(q, W),
            G = Math.min(C, $);
          (G > V && (M += A(V, G) / L(W, $, z)),
            (W = $),
            ($ = v(x, ((++D + 1) * 12) / z)));
        }
        ((U = ((F * N) / z) * M),
          (i = `Accrued coupon interest: quasi-coupon periods anchored to first_interest, basis ${T}, frequency ${z}; calc_method ${j} starts at ${!j && C > x ? "first_interest" : "issue"}. No business-day adjustment or automatic rounding.`));
      }
    }
    return (
      (U = u(U)),
      r ||
        l({
          status: "calculated",
          message:
            "Calculated under the displayed sign, timing and day-count convention.",
        }),
      U
    );
  } catch (E) {
    throw (
      r ||
        l({
          status: "error",
          error: n.isIssue(E) ? (E.error ?? "#VALUE!") : "#VALUE!",
          message:
            "An input or dependent calculation failed; inspect the source values and dates.",
        }),
      E
    );
  }
}
var Ca = ["Calibri", "Aptos", "Arial", "Times New Roman", "Courier New"],
  Ne = {
    maxRows: 1048576,
    maxCols: 16384,
    maxCells: 1e5,
    maxFileBytes: 10485760,
    maxNativeBytes: 52428800,
    maxLLMCells: 500,
    maxPatches: 500,
  };
var dr = {
  ...Ne,
  maxCellTextLength: 32767,
  maxFormulaLength: 4096,
  maxFormulaRangeCells: 2e4,
  maxFormulaDepth: 64,
  maxSpillCells: 1e4,
  maxXlsxExpandedBytes: 50 * 1024 * 1024,
  maxXlsxEntries: 2e3,
  maxXlsxExpansionRatio: 200,
  xlsxRatioThresholdBytes: 1024 * 1024,
  maxXlsxCellRecords: Ne.maxCells,
};
function Sp(e, t, n) {
  if (!e.some(Array.isArray)) return t(e);
  let i = e.map((a) =>
      Array.isArray(a) ? (a.every(Array.isArray) ? a : [a]) : [[a]],
    ),
    r = Math.max(...i.map((a) => a.length)),
    o = Math.max(...i.map((a) => a[0]?.length ?? 0));
  return !r ||
    !o ||
    i.some(
      (a) =>
        !a.length ||
        !a[0].length ||
        a.some((s) => s.length !== a[0].length) ||
        (a.length !== r && a.length !== 1) ||
        (a[0].length !== o && a[0].length !== 1),
    )
    ? n("#VALUE!")
    : r * o > dr.maxSpillCells
      ? n("#LIMIT!")
      : Array.from({ length: r }, (a, s) =>
          Array.from({ length: o }, (l, c) =>
            t(
              i.map(
                (u) => u[u.length === 1 ? 0 : s][u[0].length === 1 ? 0 : c],
              ),
            ),
          ),
        );
}
var LP = [
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
  Cp = Symbol("not-text-function"),
  Aw = (e) => (Array.isArray(e) ? e.flatMap(Aw) : [e]),
  Gi = (e) =>
    e == null ? "" : typeof e == "boolean" ? (e ? "TRUE" : "FALSE") : String(e);
function UP(e) {
  let t =
      /^(?:(\d{4})-(\d{1,2})-(\d{1,2})|(\d{1,2})\/(\d{1,2})\/(\d{4}))(?:[ T](.+))?$/.exec(
        e,
      ),
    n = 0,
    i = e;
  if (t) {
    let c = Number(t[1] ?? t[6]),
      u = Number(t[2] ?? t[4]),
      d = Number(t[3] ?? t[5]);
    if (c < 1900 || c > 9999 || u < 1 || u > 12 || d < 1) return;
    let f = c === 1900 && u === 2 && d === 29,
      p = Date.UTC(c, u - 1, d),
      h = new Date(p);
    if (!f && (h.getUTCMonth() !== u - 1 || h.getUTCDate() !== d)) return;
    if (
      ((n = f
        ? 60
        : (p - Date.UTC(1899, 11, 31)) / 864e5 +
          (p >= Date.UTC(1900, 2, 1) ? 1 : 0)),
      !t[7])
    )
      return n;
    i = t[7];
  }
  let r = /^(\d{1,2}):(\d{2})(?::(\d{2}(?:\.\d+)?))?\s*(AM|PM)?$/i.exec(i);
  if (!r) return;
  let o = Number(r[1]),
    a = Number(r[2]),
    s = Number(r[3] ?? 0),
    l = r[4]?.toUpperCase();
  if (!(a > 59 || s >= 60 || (l ? o < 1 || o > 12 : o > 23)))
    return (
      l && (o = (o % 12) + (l === "PM" ? 12 : 0)),
      n + (o * 3600 + a * 60 + s) / 86400
    );
}
function Iw(e, t, n) {
  if (!LP.includes(e)) return Cp;
  let i = (m = "#VALUE!") => {
      throw n.issue(m);
    },
    r = {
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
    [o, a] = r[e];
  (t.length < o || t.length > a) && i();
  let s = t.map(n.retrieve),
    l = s.flatMap(Aw).find(n.isIssue);
  if (l) return l;
  ((!(e === "CONCAT" || e === "TEXTJOIN") && s.some(Array.isArray)) ||
    (e === "TEXTJOIN" && s.slice(0, 2).some(Array.isArray))) &&
    i();
  let u = (m) => (m.length <= 32767 ? m : i("#LIMIT!")),
    d = (m, g) => {
      if (m === undefined && g !== undefined) return g;
      if (Array.isArray(m) || (typeof m == "string" && !m.trim())) return i();
      let b = Number(m ?? 0);
      return Number.isFinite(b) ? b : i();
    },
    f = (m, g) => Math.trunc(d(m, g)),
    p = (m) =>
      typeof m == "string"
        ? (/^(true|false)$/i.test(m) || i(), m.toLowerCase() === "true")
        : !!m,
    h = Gi(s[0]);
  if (e === "LEN") return [...h].length;
  if (e === "TRIM") return u(h.replace(/^ +| +$/g, "").replace(/ +/g, " "));
  if (e === "CLEAN") return u(h.replace(/[\u0000-\u001f]/g, ""));
  if (e === "UPPER") return u(h.toUpperCase());
  if (e === "LOWER") return u(h.toLowerCase());
  if (e === "PROPER")
    return u(
      h
        .toLowerCase()
        .replace(
          /(^|[^\p{L}\p{M}])(\p{L})/gu,
          (m, g, b) => g + b.toUpperCase(),
        ),
    );
  if (e === "EXACT") return h === Gi(s[1]);
  if (e === "LEFT" || e === "RIGHT") {
    let m = f(s[1], 1),
      g = [...h];
    return (
      m < 0 && i(),
      (e === "LEFT" ? g.slice(0, m) : m ? g.slice(-m) : []).join("")
    );
  }
  if (e === "MID" || e === "REPLACE") {
    let m = f(s[1]),
      g = f(s[2]),
      b = [...h];
    return (
      (m < 1 || g < 0) && i(),
      e === "MID"
        ? b.slice(m - 1, m - 1 + g).join("")
        : u(b.slice(0, m - 1).join("") + Gi(s[3]) + b.slice(m - 1 + g).join(""))
    );
  }
  if (e === "FIND" || e === "SEARCH") {
    let m = [...Gi(s[1])],
      g = f(s[2], 1);
    (g < 1 || g > Math.max(1, m.length)) && i();
    let b = [...h],
      w = [];
    for (let O = 0; O < b.length; O++) {
      let E = b[O];
      e === "SEARCH" && E === "~" && ["*", "?", "~"].includes(b[O + 1])
        ? w.push({ char: b[++O], kind: "literal" })
        : w.push({
            char: E,
            kind:
              e === "SEARCH" && E === "*"
                ? "many"
                : e === "SEARCH" && E === "?"
                  ? "one"
                  : "literal",
          });
    }
    w.length * m.length > 2e6 && i("#LIMIT!");
    let _ = Array(m.length + 1).fill(true);
    for (let O = w.length - 1; O >= 0; O--) {
      let E = w[O],
        R = Array(m.length + 1).fill(false);
      R[m.length] = E.kind === "many" && _[m.length];
      for (let I = m.length - 1; I >= g - 1; I--)
        R[I] =
          E.kind === "many"
            ? _[I] || R[I + 1]
            : _[I + 1] &&
              (E.kind === "one" ||
                (e === "SEARCH"
                  ? m[I].toLowerCase() === E.char.toLowerCase()
                  : m[I] === E.char));
      _ = R;
    }
    let S = _.findIndex((O, E) => E >= g - 1 && O);
    return S < 0 ? i() : S + 1;
  }
  if (e === "SUBSTITUTE") {
    let m = Gi(s[1]),
      g = Gi(s[2]),
      b = s[3] === undefined ? undefined : f(s[3]);
    if ((b !== undefined && b < 1 && i(), !m)) return h;
    let w = h.split(m),
      _ = w.length - 1;
    return (
      h.length +
        (b === undefined ? _ : b <= _ ? 1 : 0) * (g.length - m.length) >
        32767 && i("#LIMIT!"),
      w.reduce(
        (S, O, E) =>
          E === 0 ? O : S + (b === undefined || b === E ? g : m) + O,
        "",
      )
    );
  }
  if (e === "CONCAT" || e === "CONCATENATE" || e === "TEXTJOIN") {
    let m = e === "TEXTJOIN" ? Gi(s[0]) : "",
      g = e === "TEXTJOIN" && p(s[1]),
      b = s
        .slice(e === "TEXTJOIN" ? 2 : 0)
        .flatMap(Aw)
        .filter((w) => !g || (w != null && w !== ""))
        .map(Gi);
    return (
      b.reduce((w, _) => w + _.length, 0) +
        Math.max(0, b.length - 1) * m.length >
        32767 && i("#LIMIT!"),
      b.join(m)
    );
  }
  if (e === "NUMBERVALUE") {
    let m = [...Gi(s[1] === undefined ? "." : s[1])][0],
      g = [...Gi(s[2] === undefined ? "," : s[2])][0];
    (!m || !g || m === g || /[0-9+\-eE%]/.test(m) || /[0-9+\-eE%]/.test(g)) &&
      i();
    let b = h.replace(/ /g, "");
    if (!b) return 0;
    let w = /%+$/.exec(b)?.[0].length ?? 0;
    b = b.slice(0, b.length - w);
    let _ = b.split(m);
    ((_.length > 2 || _[1]?.includes(g)) && i(),
      (b = _[0].split(g).join("") + (_.length === 2 ? "." + _[1] : "")),
      /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(b) || i());
    let S = Number(b);
    Number.isFinite(S) || i("#NUM!");
    for (let O = 0; O < w; O++) S /= 100;
    return S;
  }
  if (e === "VALUE") {
    let m = n.temporal ? n.temporal(h.trim()) : UP(h.trim());
    if (m !== undefined) return m;
    let g = h.trim(),
      b = false;
    g.startsWith("(") &&
      g.endsWith(")") &&
      ((b = true), (g = g.slice(1, -1).trim()));
    let w = g.endsWith("%");
    (w && (g = g.slice(0, -1).trim()),
      (g = g.replace(/^\$\s*/, "")),
      b && /^[+-]/.test(g) && i(),
      /^[+-]?(?:(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(
        g,
      ) || i());
    let _ = (Number(g.replace(/,/g, "")) * (b ? -1 : 1)) / (w ? 100 : 1);
    return Number.isFinite(_) ? _ : i("#NUM!");
  }
  if (e === "TEXT") {
    ((typeof s[1] != "string" || s[1].length > 256) && i(), d(s[0]));
    let m = n.format();
    return typeof m == "string" ? u(m) : m;
  }
  if (e === "TEXTBEFORE" || e === "TEXTAFTER") {
    let m = Gi(s[1]),
      g = f(s[2], 1),
      b = d(s[3], 0),
      w = d(s[4], 0);
    if (((!g || ![0, 1].includes(b) || ![0, 1].includes(w)) && i(), !h))
      return "";
    if ((Math.abs(g) > [...h].length && i(), !m))
      return (e === "TEXTBEFORE") == g > 0 ? "" : h;
    let _ = m.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"),
      S = new RegExp(g < 0 ? `(?=${_})` : _, b ? "giu" : "gu"),
      O = [...h.matchAll(S)].map((R) => R.index);
    if (g < 0) {
      let R = [],
        I = h.length;
      for (let U = O.length - 1; U >= 0; U--)
        O[U] + m.length <= I && (R.push(O[U]), (I = O[U]));
      O = R.reverse();
    }
    let E = O[g > 0 ? g - 1 : O.length + g];
    return E === undefined
      ? w && Math.abs(g) === O.length + 1
        ? (e === "TEXTBEFORE") == g > 0
          ? h
          : ""
        : s[5] !== undefined
          ? s[5]
          : i("#N/A")
      : e === "TEXTBEFORE"
        ? h.slice(0, E)
        : h.slice(E + m.length);
  }
  return Cp;
}
var Ts = ["top", "right", "bottom", "left"],
  Nw = ["none", "thin", "medium", "thick", "dotted", "dashed", "double"];
var FP = new Set([
  "binding",
  "bold",
  "italic",
  "underline",
  "strikethrough",
  "fontFamily",
  "fontSize",
  "wrap",
  "border",
  "borders",
  "fill",
  "color",
  "align",
  "verticalAlign",
  "indent",
]);
function fn(e) {
  throw Error(e);
}
function ju(e) {
  return !!(
    e &&
    typeof e == "object" &&
    !Array.isArray(e) &&
    [Object.prototype, null].includes(Object.getPrototypeOf(e))
  );
}
function Rw(e) {
  return (
    (typeof e != "string" || !/^#[0-9a-f]{3}([0-9a-f]{3})?$/i.test(e)) &&
      fn("Style colors must use hexadecimal colors."),
    e
  );
}
function zP(e) {
  return (
    (!ju(e) ||
      Object.keys(e).some((t) => !["style", "color"].includes(t)) ||
      !Nw.includes(e.style)) &&
      fn("Choose a supported border line style and hexadecimal color."),
    e.color !== undefined && Rw(e.color),
    e.style === "none"
      ? { style: "none" }
      : {
          style: e.style,
          color: e.color === undefined ? "#000000" : Rw(e.color),
        }
  );
}
function As(e) {
  ju(e) || fn("Cell style must be an object.");
  for (let n of Object.keys(e))
    FP.has(n) || fn(`Unsupported cell style setting: ${n}.`);
  let t = {};
  e.binding !== undefined && (t.binding = Ow(e.binding));
  for (let n of ["bold", "italic", "underline", "strikethrough", "wrap"])
    e[n] !== undefined &&
      (typeof e[n] != "boolean" && fn(`Style ${n} must be boolean.`),
      (t[n] = e[n]));
  (e.fontFamily !== undefined &&
    (Ca.includes(e.fontFamily) || fn("Unsupported font family."),
    (t.fontFamily = e.fontFamily)),
    e.fontSize !== undefined &&
      ((typeof e.fontSize != "number" ||
        !Number.isInteger(e.fontSize) ||
        e.fontSize < 8 ||
        e.fontSize > 72) &&
        fn("Font size must be an integer between 8 and 72."),
      (t.fontSize = e.fontSize)));
  for (let n of ["fill", "color"]) e[n] !== undefined && (t[n] = Rw(e[n]));
  if (
    (e.align !== undefined &&
      (["left", "center", "right"].includes(e.align) ||
        fn("Invalid text alignment."),
      (t.align = e.align)),
    e.verticalAlign !== undefined &&
      (["top", "middle", "bottom"].includes(e.verticalAlign) ||
        fn("Choose Top, Center or Bottom vertical alignment."),
      (t.verticalAlign = e.verticalAlign)),
    e.indent !== undefined &&
      ((typeof e.indent != "number" ||
        !Number.isInteger(e.indent) ||
        e.indent < 0 ||
        e.indent > 15) &&
        fn("Indent must be an integer between 0 and 15."),
      (t.indent = e.indent)),
    e.border !== undefined &&
      (["all", "bottom", "none"].includes(e.border) ||
        fn("Unsupported cell border."),
      (t.border = e.border)),
    e.borders !== undefined)
  ) {
    (!ju(e.borders) || Object.keys(e.borders).some((i) => !Ts.includes(i))) &&
      fn("Cell borders support top, right, bottom and left edges only.");
    let n = {};
    for (let i of Ts) e.borders[i] !== undefined && (n[i] = zP(e.borders[i]));
    Object.keys(n).length && (t.borders = n);
  }
  return t;
}
function Ow(e) {
  (!ju(e) ||
    Object.keys(e).some(
      (n) =>
        !["namespace", "styleId", "base", "missing", "themeFallback"].includes(
          n,
        ),
    )) &&
    fn(
      "Named style binding must contain a namespace, style ID and supported base.",
    );
  for (let n of ["namespace", "styleId"])
    (typeof e[n] != "string" ||
      !e[n].trim() ||
      e[n].length > 100 ||
      /[\x00-\x1f\x7f]/.test(e[n])) &&
      fn(
        "Named style identifiers must contain 1\u2013100 characters without control characters.",
      );
  (!ju(e.base) || Object.hasOwn(e.base, "binding")) &&
    fn("A named style base cannot contain another binding.");
  for (let n of ["missing", "themeFallback"])
    e[n] !== undefined &&
      e[n] !== true &&
      fn("Invalid named style fallback flag.");
  let t = As(e.base);
  return (
    e.missing &&
      Object.keys(t).length &&
      fn("A missing named style must use an empty fallback base."),
    {
      namespace: e.namespace,
      styleId: e.styleId,
      base: t,
      ...(e.missing ? { missing: true } : {}),
      ...(e.themeFallback ? { themeFallback: true } : {}),
    }
  );
}
function Mw(e) {
  if (!e) return {};
  let { binding, ...n } = e;
  if (!binding) return n;
  let i = binding.missing ? {} : binding.base,
    r = { ...i, ...n };
  return (
    Object.hasOwn(n, "border")
      ? delete r.borders
      : i.borders && (r.borders = { ...i.borders }),
    n.borders && (r.borders = { ...r.borders, ...n.borders }),
    r
  );
}
function Wu(e, t) {
  if (((e = Mw(e)), e?.borders?.[t])) return e.borders[t];
  if (e?.border === "none") return { style: "none" };
  if (e?.border === "all" || (e?.border === "bottom" && t === "bottom"))
    return { style: "thin", color: "#000000" };
}
var Ap = [
    "background",
    "text",
    "accent1",
    "accent2",
    "accent3",
    "accent4",
    "accent5",
    "accent6",
  ],
  BP = { maxStyles: 50, maxCells: 500, maxName: 80 },
  U6 = {
    name: "Office",
    palette: {
      background: "#FFFFFF",
      text: "#000000",
      accent1: "#4472C4",
      accent2: "#ED7D31",
      accent3: "#A5A5A5",
      accent4: "#FFC000",
      accent5: "#5B9BD5",
      accent6: "#70AD47",
    },
    fonts: { body: "Calibri", heading: "Calibri" },
  };
function Fr(e) {
  throw Error(e);
}
function bo(e) {
  return !!(
    e &&
    typeof e == "object" &&
    !Array.isArray(e) &&
    [Object.prototype, null].includes(Object.getPrototypeOf(e))
  );
}
function Ta(e, t, n) {
  Object.keys(e).some((i) => !t.includes(i)) && Fr(`Unsupported ${n} setting.`);
}
function Tp(e, t, n = 80) {
  return (
    (typeof e != "string" ||
      !e.trim() ||
      e.length > n ||
      /[\x00-\x1f\x7f]/.test(e)) &&
      Fr(
        `${t} must contain 1\u2013${n} characters without control characters.`,
      ),
    e.trim()
  );
}
function F6(e) {
  return structuredClone(e);
}
function Gl(e) {
  return Array.isArray(e)
    ? `[${e.map(Gl).join(",")}]`
    : bo(e)
      ? `{${Object.keys(e)
          .filter((t) => e[t] !== undefined)
          .sort()
          .map((t) => `${JSON.stringify(t)}:${Gl(e[t])}`)
          .join(",")}}`
      : JSON.stringify(e);
}
function jP(e) {
  (bo(e) || Fr("Choose a workbook theme."),
    Ta(e, ["name", "palette", "fonts"], "theme"),
    (!bo(e.palette) || !bo(e.fonts)) &&
      Fr("A theme needs eight colors and body and heading fonts."),
    Ta(e.palette, Ap, "theme palette"),
    Ta(e.fonts, ["body", "heading"], "theme font"));
  let t = {};
  for (let n of Ap) {
    let i = e.palette[n];
    ((typeof i != "string" || !/^#[0-9a-f]{6}$/i.test(i)) &&
      Fr("Theme colors must use six-digit hexadecimal colors."),
      (t[n] = i.toUpperCase()));
  }
  for (let n of ["body", "heading"])
    Ca.includes(e.fonts[n]) || Fr("Choose a supported theme font.");
  return {
    name: Tp(e.name, "Theme name"),
    palette: t,
    fonts: { body: e.fonts.body, heading: e.fonts.heading },
  };
}
function WP(e) {
  (bo(e) || Fr("Named style definition must be an object."),
    Ta(
      e,
      ["base", "fontRole", "fillRole", "colorRole", "borderRoles"],
      "named style",
    ),
    (!bo(e.base) || Object.hasOwn(e.base, "binding")) &&
      Fr(
        "A named style definition needs supported formatting without another binding.",
      ));
  let t = { base: As(e.base) };
  e.fontRole !== undefined &&
    (["body", "heading"].includes(e.fontRole) ||
      Fr("Choose Body or Heading theme font."),
    (t.fontRole = e.fontRole));
  for (let n of ["fillRole", "colorRole"])
    e[n] !== undefined &&
      (Ap.includes(e[n]) || Fr("Choose a supported theme color role."),
      (t[n] = e[n]));
  if (e.borderRoles !== undefined) {
    (bo(e.borderRoles) || Fr("Border color roles must be an object."),
      Ta(e.borderRoles, Ts, "border role"));
    let n = {};
    for (let i of Ts)
      e.borderRoles[i] !== undefined &&
        (Ap.includes(e.borderRoles[i]) ||
          Fr("Choose a supported border color role."),
        (!Wu(t.base, i) || Wu(t.base, i)?.style === "none") &&
          Fr("Add a border line before assigning its theme color."),
        (n[i] = e.borderRoles[i]));
    Object.keys(n).length && (t.borderRoles = n);
  }
  return t;
}
function VP(e) {
  (bo(e) || Fr("Workbook styles must be an object."),
    Ta(
      e,
      ["namespace", "theme", "styles", "missingTheme"],
      "workbook style catalog",
    ),
    e.missingTheme !== undefined &&
      e.missingTheme !== true &&
      Fr("Invalid missing theme flag."),
    (!Array.isArray(e.styles) || e.styles.length > BP.maxStyles) &&
      Fr("A workbook supports at most 50 named styles."));
  let t = new Set(),
    n = new Set(),
    i = e.styles.map((a) => {
      (bo(a) || Fr("Invalid named style."),
        Ta(a, ["id", "name", "definition"], "named style"));
      let s = Tp(a.id, "Style ID", 100),
        l = Tp(a.name, "Style name");
      return (
        (t.has(s) || n.has(l.toLocaleLowerCase("en-US"))) &&
          Fr("Named style IDs and names must be unique within this workbook."),
        t.add(s),
        n.add(l.toLocaleLowerCase("en-US")),
        { id: s, name: l, definition: WP(a.definition) }
      );
    }),
    r = e.theme === undefined || e.missingTheme === true,
    o = e.theme === undefined ? F6(U6) : jP(e.theme);
  return (
    r &&
      Gl(o) !== Gl(U6) &&
      Fr("A missing theme must use the Office fallback."),
    {
      namespace: Tp(e.namespace, "Style namespace", 100),
      theme: o,
      styles: i,
      ...(r ? { missingTheme: true } : {}),
    }
  );
}
function qP(e, t) {
  let n = F6(e.base);
  (e.fontRole && (n.fontFamily = t.fonts[e.fontRole]),
    e.fillRole && (n.fill = t.palette[e.fillRole]),
    e.colorRole && (n.color = t.palette[e.colorRole]));
  for (let i of Ts)
    e.borderRoles?.[i] &&
      (n.borders = {
        ...n.borders,
        [i]: { ...Wu(n, i), color: t.palette[e.borderRoles[i]] },
      });
  return As(n);
}
function XP(e, t) {
  e &&
    t.namespace !== e.namespace &&
    Fr(
      "This style belongs to another workbook. Import or remap its style catalog before pasting the binding.",
    );
  let n = e?.styles.find((i) => i.id === t.styleId);
  return {
    namespace: t.namespace,
    styleId: t.styleId,
    base: n && e ? qP(n.definition, e.theme) : {},
    ...(n ? {} : { missing: true }),
    ...(e?.missingTheme ? { themeFallback: true } : {}),
  };
}
function z6(e, t) {
  let n = e === undefined ? undefined : VP(e);
  n ? (t.styleCatalog = n) : delete t.styleCatalog;
  let i = bo(e) && e.theme === undefined;
  for (let r of t.sheets)
    for (let o of cf(r.cells) ? Object.values(r.cells) : []) {
      if (!o.style?.binding) continue;
      let a = Ow(o.style.binding),
        s = XP(n, a);
      (!s.missing &&
        !i &&
        Gl(a) !== Gl(s) &&
        Fr(
          "Named style appearance does not match the workbook catalog. Reapply the style from a trusted catalog.",
        ),
        (o.style.binding = s));
    }
}
function HP(e) {
  if (e?.binding?.missing)
    return "Named style is missing. Default formatting and direct cell formatting are shown.";
  if (e?.binding?.themeFallback)
    return "Workbook theme is missing. The Office theme is shown.";
}
function B6(e) {
  let t = new Set();
  e.styleCatalog?.missingTheme &&
    t.add("Workbook theme is missing. The Office theme is shown.");
  for (let n of e.sheets)
    for (let i of Object.values(n.cells)) {
      let r = HP(i.style);
      r && t.add(r);
    }
  return [...t];
}
var Bo = {
    threads: 1e3,
    messagesPerThread: 100,
    totalMessages: 5e3,
    textLength: 4e3,
    authorLength: 100,
    urlLength: 2048,
    labelLength: 256,
  },
  $w =
    /[\u0000-\u001f\u007f-\u009f\u200b-\u200f\u202a-\u202e\u2060-\u206f\ufeff]/,
  Er = (e) => {
    throw new Error(e);
  };
function Aa(e, t) {
  return (
    (!e ||
      typeof e != "object" ||
      Array.isArray(e) ||
      ![Object.prototype, null].includes(Object.getPrototypeOf(e))) &&
      Er(`${t} must be a plain object.`),
    e
  );
}
function Zi(e, t, n, i = false) {
  return (
    (typeof e != "string" ||
      !e.trim() ||
      e.length > n ||
      $w.test(i ? e.replace(/[\n\r\t]/g, "") : e)) &&
      Er(
        `${t} must contain 1\u2013${n} characters without unsupported control characters.`,
      ),
    e.trim()
  );
}
function Vu(e, t = "Review timestamp") {
  return (
    (typeof e != "string" ||
      !Number.isFinite(Date.parse(e)) ||
      new Date(e).toISOString() !== e) &&
      Er(`${t} must be an ISO UTC timestamp.`),
    e
  );
}
function W6(e) {
  if (typeof e != "string") return Er("Choose a valid worksheet cell address.");
  let t = tt(e.trim());
  return `${Ue(t.col)}${t.row + 1}`;
}
function j6(e = "Reviewer") {
  return Zi(e, "Reviewer name", Bo.authorLength);
}
function Pw(e) {
  let t = Aa(e, "Hyperlink"),
    n =
      t.label === undefined || t.label === ""
        ? undefined
        : Zi(t.label, "Link display text", Bo.labelLength);
  if (t.kind === "internal")
    return {
      kind: "internal",
      sheetId: Zi(t.sheetId, "Link worksheet ID", 100),
      address: W6(t.address),
      ...(n ? { label: n } : {}),
    };
  if (t.kind !== "external")
    return Er("Choose a web, email or worksheet link.");
  let i = Zi(t.url, "Link address", Bo.urlLength),
    r = i;
  for (let a = 0; a < 4; a++) {
    ($w.test(r) || /\\/.test(r)) &&
      Er("Link addresses cannot contain control characters or backslashes.");
    let s = r;
    try {
      s = decodeURIComponent(r);
    } catch {
      Er("The link contains invalid encoded characters.");
    }
    if (s === r) break;
    r = s;
  }
  ($w.test(r) || /%(?:0[\da-f]|1[\da-f]|7f|8[\da-f]|9[\da-f])/i.test(r)) &&
    Er("Link addresses cannot contain encoded control characters.");
  let o;
  try {
    o = new URL(i);
  } catch {
    return Er("Enter a complete https://, http:// or mailto: address.");
  }
  return (
    ["https:", "http:", "mailto:"].includes(o.protocol) ||
      Er("Only https://, http:// and mailto: links are supported."),
    (o.username || o.password) &&
      Er("Links cannot contain embedded credentials."),
    o.protocol === "mailto:"
      ? (o.search ||
          o.hash ||
          o.host ||
          !/^[^\s@/?#:,]+@[^\s@/?#:,]+\.[^\s@/?#:,]+$/.test(
            decodeURIComponent(o.pathname),
          )) &&
        Er("Use mailto: with one email address and no additional headers.")
      : o.hostname || Er("Web links require a hostname."),
    o.href.length > Bo.urlLength && Er("The normalized link is too long."),
    { kind: "external", url: o.href, ...(n ? { label: n } : {}) }
  );
}
function V6(e, t) {
  let n = Aa(e, "Workbook review metadata");
  if (n.reviewThreads !== undefined) {
    (!Array.isArray(n.reviewThreads) || n.reviewThreads.length > Bo.threads) &&
      Er(`A workbook can contain at most ${Bo.threads} review threads.`);
    let i = new Set(),
      r = new Set(),
      o = 0,
      a = n.reviewThreads.map((s) => {
        let l = Aa(s, "Review thread"),
          c = Aa(l.anchor, "Review target"),
          u = Zi(l.id, "Review thread ID", 100);
        (i.has(u) && Er("Review thread IDs must be unique."), i.add(u));
        let d = Vu(l.createdAt),
          f = Vu(l.updatedAt);
        f < d && Er("Review update time cannot precede creation.");
        let p = {
          sheetId: Zi(c.sheetId, "Review worksheet ID", 100),
          sheetName: Zi(c.sheetName, "Review worksheet name", 31),
          address: W6(c.address),
        };
        (!Array.isArray(l.messages) ||
          !l.messages.length ||
          l.messages.length > Bo.messagesPerThread ||
          (o += l.messages.length) > Bo.totalMessages) &&
          Er(
            "Review discussions exceed the per-thread or workbook message limit.",
          );
        let h = d,
          m = l.messages.map((b, w) => {
            let _ = Aa(b, "Review message"),
              S = Zi(_.id, "Review message ID", 100);
            (r.has(S) &&
              Er("Review message IDs must be unique within the workbook."),
              r.add(S));
            let O = Vu(_.createdAt);
            return (
              (O < h || O > f || (w === 0 && O !== d)) &&
                Er(
                  "Review message times must follow their discussion in order.",
                ),
              (h = O),
              {
                id: S,
                author: j6(Zi(_.author, "Reviewer name", 100)),
                text: Zi(_.text, "Comment", Bo.textLength, true),
                createdAt: O,
              }
            );
          }),
          g = { id: u, anchor: p, messages: m, createdAt: d, updatedAt: f };
        if (l.resolved !== undefined) {
          let b = Aa(l.resolved, "Review resolution"),
            w = Vu(b.createdAt);
          ((w < m.at(-1).createdAt || w > f) &&
            Er("Invalid review resolution time."),
            (g.resolved = {
              author: j6(Zi(b.author, "Reviewer name", 100)),
              createdAt: w,
            }));
        }
        if (l.orphaned !== undefined) {
          let b = Aa(l.orphaned, "Orphaned review target"),
            w = Vu(b.createdAt);
          ((!["worksheet-deleted", "cell-deleted"].includes(String(b.reason)) ||
            w < d ||
            w > f) &&
            Er("Invalid orphaned review target."),
            (g.orphaned = { reason: b.reason, createdAt: w }));
        }
        return (
          !g.orphaned &&
            !t.sheets.some((b) => b.id === p.sheetId) &&
            Er(
              "A review thread with a missing worksheet must be retained as orphaned.",
            ),
          g
        );
      });
    a.length ? (t.reviewThreads = a) : delete t.reviewThreads;
  }
  for (let i of t.sheets)
    for (let r of cf(i.cells) ? Object.values(i.cells) : []) {
      if (r.hyperlink === undefined) continue;
      let o = Pw(r.hyperlink);
      (o.kind === "internal" &&
        !t.sheets.some((a) => a.id === o.sheetId) &&
        Er("An internal hyperlink must target an existing worksheet."),
        (r.hyperlink = o));
    }
}
function q6(e, t) {
  if (e === undefined) return;
  if (
    !e ||
    typeof e != "object" ||
    Array.isArray(e) ||
    (Object.getPrototypeOf(e) !== Object.prototype &&
      Object.getPrototypeOf(e) !== null) ||
    Object.keys(e).some(
      (i) => !["horizontal", "vertical", "row", "col"].includes(i),
    )
  )
    throw Error(
      "Split panes require divider percentages and valid row/column anchors.",
    );
  let n = e;
  if (n.horizontal === undefined && n.vertical === undefined)
    throw Error("Choose a horizontal split, vertical split, or both.");
  for (let i of [n.horizontal, n.vertical])
    if (i !== undefined && (!Number.isFinite(i) || i < 20 || i > 80))
      throw Error("Split dividers must be between 20% and 80%.");
  if (
    !Number.isInteger(n.row) ||
    n.row < 0 ||
    n.row >= t.rows ||
    !Number.isInteger(n.col) ||
    n.col < 0 ||
    n.col >= t.cols
  )
    throw Error(
      "Split anchors must be valid zero-based worksheet coordinates.",
    );
  return {
    row: n.row,
    col: n.col,
    ...(n.horizontal !== undefined ? { horizontal: n.horizontal } : {}),
    ...(n.vertical !== undefined ? { vertical: n.vertical } : {}),
  };
}
function X6(e, t) {
  let n = t === "row" ? 21 : 40,
    i = t === "row" ? 546 : 600;
  if (typeof e != "number" || !Number.isFinite(e) || e < n || e > i)
    throw Error(
      (t === "row" ? "Default row height" : "Default column width") +
        " must be " +
        n +
        "\u2013" +
        i +
        " pixels.",
    );
  return e;
}
function H6(e, t) {
  let n =
      e.defaultRowHeight === undefined
        ? undefined
        : X6(e.defaultRowHeight, "row"),
    i =
      e.defaultColumnWidth === undefined
        ? undefined
        : X6(e.defaultColumnWidth, "column");
  (n !== undefined && (t.defaultRowHeight = n),
    i !== undefined && (t.defaultColumnWidth = i));
}
var Na = {
  maxIntervalsPerAxis: 1e4,
  maxFrozenVisibleRows: 100,
  maxFrozenVisibleCols: 50,
  maxXlsxHiddenRows: 1e5,
};
function ZP(e) {
  if (!Number.isSafeInteger(e) || e < 0)
    throw Error("Axis capacity must be a nonnegative integer.");
}
function Vo(e, t) {
  if ((ZP(t), e === undefined)) return [];
  if (!Array.isArray(e) || e.length > Na.maxIntervalsPerAxis)
    throw Error("Hidden indices require at most 10,000 intervals per axis.");
  let n = e
      .map((r) => {
        if (
          !r ||
          typeof r != "object" ||
          Array.isArray(r) ||
          (Object.getPrototypeOf(r) !== Object.prototype &&
            Object.getPrototypeOf(r) !== null) ||
          Object.keys(r).some((o) => o !== "start" && o !== "end") ||
          !Number.isInteger(r.start) ||
          !Number.isInteger(r.end) ||
          r.start < 0 ||
          r.end < r.start ||
          r.end >= t
        )
          throw Error(
            "Hidden intervals require valid zero-based inclusive axis bounds.",
          );
        return { start: r.start, end: r.end };
      })
      .sort((r, o) => r.start - o.start || r.end - o.end),
    i = [];
  for (let r of n) {
    let o = i.at(-1);
    o && r.start <= o.end + 1 ? (o.end = Math.max(o.end, r.end)) : i.push(r);
  }
  return i;
}
function Oa(e, t = []) {
  let n = Vo(t, e),
    i = [0];
  for (let c of n) i.push(i.at(-1) + c.end - c.start + 1);
  let r = e - i.at(-1),
    o = (c) => {
      let u = 0,
        d = n.length;
      for (; u < d;) {
        let f = (u + d) >>> 1;
        n[f].start <= c ? (u = f + 1) : (d = f);
      }
      return u - 1;
    },
    a = (c) => {
      let u = Math.max(0, Math.min(e, Math.floor(c))),
        d = o(u - 1),
        f =
          d < 0
            ? 0
            : i[d] + Math.min(u - n[d].start, n[d].end - n[d].start + 1);
      return u - f;
    },
    s = (c) => {
      let u = o(c);
      return u >= 0 && c <= n[u].end;
    },
    l = (c) => {
      if (!Number.isInteger(c) || c < 0 || c >= r) return -1;
      let u = 0,
        d = e - 1;
      for (; u < d;) {
        let f = Math.floor((u + d) / 2);
        a(f + 1) <= c ? (u = f + 1) : (d = f);
      }
      return u;
    };
  return {
    count: r,
    logicalCount: e,
    logicalIndex: l,
    countBefore: a,
    isHidden: s,
    viewIndex: (c) =>
      Number.isInteger(c) && c >= 0 && c < e && !s(c) ? a(c) : null,
    next: (c, u, d = false) => {
      let f = Math.trunc(c) + (d ? 0 : u);
      if (!r || (u === 1 && f >= e) || (u === -1 && f < 0)) return null;
      let p = u === 1 ? a(Math.max(0, f)) : a(Math.min(e, f + 1)) - 1,
        h = l(p);
      return h < 0 ? null : h;
    },
  };
}
function Dw(e, t, n = { rows: t.rowCount, cols: t.colCount }) {
  H6(e, t);
  let i = q6(e.splitPanes, n);
  if (i && (e.frozenRows || e.frozenCols))
    throw Error("Split panes and frozen panes cannot be enabled together.");
  i && (t.splitPanes = i);
  let r = Vo(e.hiddenRows, n.rows),
    o = Vo(e.hiddenCols, n.cols),
    a = { ...t, hiddenRows: r, hiddenCols: o },
    s = Oa(t.rowCount, qo(a, "row")),
    l = Oa(t.colCount, qo(a, "column"));
  if (!s.count || !l.count)
    throw Error("At least one row and one column must remain visible.");
  let c = e.frozenRows === undefined ? 0 : e.frozenRows,
    u = e.frozenCols === undefined ? 0 : e.frozenCols;
  if (
    !Number.isInteger(c) ||
    !Number.isInteger(u) ||
    c < 0 ||
    u < 0 ||
    c > n.rows ||
    u > n.cols
  )
    throw Error(
      "Frozen boundaries must be valid leading row and column counts.",
    );
  let d = s.countBefore(c),
    f = l.countBefore(u);
  if (d > Na.maxFrozenVisibleRows || f > Na.maxFrozenVisibleCols)
    throw Error(
      "Freeze panes supports at most 100 visible rows and 50 visible columns.",
    );
  if (d === s.count || f === l.count)
    throw Error("Freeze panes must leave an unfrozen visible row and column.");
  for (let p of t.merges ?? []) {
    let h = Qe(p);
    if (
      (h.firstRow < c && h.lastRow >= c) ||
      (h.firstCol < u && h.lastCol >= u)
    )
      throw Error(
        "Freeze panes cannot cross a merged range. Unmerge those cells first.",
      );
  }
  (e.hiddenRows !== undefined && (t.hiddenRows = r),
    e.hiddenCols !== undefined && (t.hiddenCols = o),
    e.frozenRows !== undefined && (t.frozenRows = c),
    e.frozenCols !== undefined && (t.frozenCols = u));
}
var Z6 = { maxGroups: 1e3, maxDepth: 8 },
  KP = (e, t) => (t === "row" ? e.rowCount : e.colCount);
function JP(e) {
  if (e !== "row" && e !== "column") throw Error("Choose rows or columns.");
}
function G6(e, t) {
  let n = e
      .filter((r) => r.axis === t)
      .sort((r, o) => r.start - o.start || o.end - r.end),
    i = [];
  return n.map((r) => {
    for (; i.length && i.at(-1).end < r.start;) i.pop();
    let o = i.at(-1);
    if (o && o.start === r.start && o.end === r.end)
      throw Error("An outline group already exists for these rows or columns.");
    if (o && o.end < r.end)
      throw Error(
        "Outline groups must be nested or separate; partially overlapping groups are not supported.",
      );
    if ((i.push(r), i.length > Z6.maxDepth))
      throw Error("Outline groups support at most 8 nested groups.");
    return { ...r, depth: i.length };
  });
}
function K6(e, t, n = { rows: t.rowCount, cols: t.colCount }) {
  if (e.outlineGroups === undefined) return;
  if (!Array.isArray(e.outlineGroups) || e.outlineGroups.length > Z6.maxGroups)
    throw Error("A worksheet supports at most 1,000 outline groups.");
  let i = new Set(),
    r = e.outlineGroups.map((o) => {
      if (
        !o ||
        typeof o != "object" ||
        Array.isArray(o) ||
        (Object.getPrototypeOf(o) !== Object.prototype &&
          Object.getPrototypeOf(o) !== null) ||
        Object.keys(o).some(
          (a) => !["id", "axis", "start", "end", "collapsed"].includes(a),
        ) ||
        typeof o.id != "string" ||
        !o.id.trim() ||
        o.id.length > 100 ||
        i.has(o.id) ||
        !["row", "column"].includes(o.axis) ||
        !Number.isSafeInteger(o.start) ||
        !Number.isSafeInteger(o.end) ||
        o.start < 0 ||
        o.end < o.start ||
        o.end >= (o.axis === "row" ? n.rows : n.cols) ||
        typeof o.collapsed != "boolean"
      )
        throw Error(
          "Outline groups require unique IDs, valid row or column bounds, and a boolean collapsed state.",
        );
      return (
        i.add(o.id),
        {
          id: o.id,
          axis: o.axis,
          start: o.start,
          end: o.end,
          collapsed: o.collapsed,
        }
      );
    });
  (G6(r, "row"), G6(r, "column"), (t.outlineGroups = r));
}
function qo(e, t) {
  JP(t);
  let n = KP(e, t),
    i = [
      ...Vo(t === "row" ? e.hiddenRows : e.hiddenCols, n),
      ...(e.outlineGroups ?? [])
        .filter((o) => o.axis === t && o.collapsed)
        .map(({ start, end }) => ({ start, end })),
    ].sort((o, a) => o.start - a.start || o.end - a.end),
    r = [];
  for (let o of i) {
    let a = r.at(-1);
    a && o.start <= a.end + 1
      ? (a.end = Math.max(a.end, o.end))
      : r.push({ ...o });
  }
  return Vo(r, n);
}
var J6 = (e) =>
  !!e &&
  typeof e == "object" &&
  !Array.isArray(e) &&
  [Object.prototype, null].includes(Object.getPrototypeOf(e));
function qu(e, t, n) {
  if (typeof e != "string" || !e.trim() || e.length > n)
    throw new Error(`${t} must contain 1\u2013${n} characters.`);
  return e.trim();
}
function Y6(e) {
  if (
    typeof e != "string" ||
    !/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}Z$/.test(e) ||
    !Number.isFinite(Date.parse(e)) ||
    new Date(e).toISOString() !== e
  )
    throw new Error("Workbook metadata requires a valid UTC timestamp.");
  return e;
}
function Lw(e, t) {
  if (e === undefined) return;
  if (
    !J6(e) ||
    Object.keys(e).some(
      (o) =>
        ![
          "createdAt",
          "modifiedAt",
          "engagement",
          "period",
          "template",
        ].includes(o),
    )
  )
    throw new Error("Invalid workbook metadata fields.");
  let n = Y6(e.createdAt),
    i = Y6(e.modifiedAt);
  if (i < n)
    throw new Error("Workbook modified time cannot precede its creation time.");
  let r = { createdAt: n, modifiedAt: i };
  if (
    (e.engagement !== undefined &&
      (r.engagement = qu(e.engagement, "Engagement", 200)),
    e.period !== undefined && (r.period = qu(e.period, "Period", 100)),
    e.template !== undefined)
  ) {
    let o = e.template;
    if (
      !J6(o) ||
      Object.keys(o).some(
        (a) =>
          ![
            "id",
            "name",
            "version",
            "sourceWorkbookId",
            "sourceRevision",
          ].includes(a),
      ) ||
      !Number.isSafeInteger(o.version) ||
      Number(o.version) < 1 ||
      (o.sourceRevision !== undefined &&
        (!Number.isSafeInteger(o.sourceRevision) ||
          Number(o.sourceRevision) < 0))
    )
      throw new Error("Invalid template provenance.");
    r.template = {
      id: qu(o.id, "Template ID", 100),
      name: qu(o.name, "Template name", 200),
      version: Number(o.version),
      sourceWorkbookId: qu(
        o.sourceWorkbookId,
        "Template source workbook ID",
        100,
      ),
      ...(o.sourceRevision !== undefined
        ? { sourceRevision: Number(o.sourceRevision) }
        : {}),
    };
  }
  t.metadata = r;
}
function Ip(e = {}, t = new Date().toISOString()) {
  let n = {};
  return (Lw({ createdAt: t, modifiedAt: t, ...e }, n), n.metadata);
}
function Q6(e = "Book1") {
  return gi({
    schemaVersion: 1,
    id: crypto.randomUUID(),
    name: e.trim(),
    revision: 0,
    metadata: Ip(),
    sheets: [
      {
        id: crypto.randomUUID(),
        name: "Sheet1",
        cells: {},
        rowCount: Ne.maxRows,
        colCount: Ne.maxCols,
        frozenRows: 0,
      },
    ],
    events: [],
  });
}
var Uw = /[A-Za-z0-9_.\\\u0080-\uffff]/,
  QP =
    /^(?:'(?:[^']|'')*'|[A-Za-z_\\\u0080-\uffff][A-Za-z0-9_.\\\u0080-\uffff]*)(?::(?:'(?:[^']|'')*'|[A-Za-z_\\\u0080-\uffff][A-Za-z0-9_.\\\u0080-\uffff]*))?!/,
  eD = (e) =>
    [...e.toUpperCase()].reduce((t, n) => t * 26 + n.charCodeAt(0) - 64, 0) - 1;
function Fw(e, t, n) {
  let i = eD(t) + (e ? 0 : n);
  return i < 0 || i >= Ne.maxCols ? "#REF!" : e + Ue(i);
}
function zw(e, t, n) {
  let i = Number(t) + (e ? 0 : n);
  return i < 1 || i > Ne.maxRows ? "#REF!" : e + i;
}
function bi(e, t, n) {
  if (!Number.isSafeInteger(t) || !Number.isSafeInteger(n))
    throw new Error("Formula offsets must be integers.");
  if (!e.startsWith("=")) return e;
  let i = "",
    r = 0;
  for (; r < e.length;) {
    let o = e.slice(r),
      a = e[r],
      s = (r === 0 || !Uw.test(e[r - 1])) && o.match(QP);
    if (s) {
      ((i += s[0]), (r += s[0].length));
      continue;
    }
    if (a === '"' || a === "'") {
      let p = r++;
      for (; r < e.length;)
        if (e[r++] === a)
          if (e[r] === a) r++;
          else break;
      i += e.slice(p, r);
      continue;
    }
    if (a === "[") {
      let p = r++,
        h = 1;
      for (; r < e.length && h;)
        (e[r] === "[" && h++, e[r] === "]" && h--, r++);
      i += e.slice(p, r);
      continue;
    }
    let l = r === 0 || !Uw.test(e[r - 1]),
      c = l ? /^(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})/.exec(o) : null,
      u = l ? /^(\$?)([1-9]\d*):(\$?)([1-9]\d*)/.exec(o) : null,
      d = l ? /^(\$?)([A-Za-z]{1,3})(\$?)([1-9]\d*)/.exec(o) : null,
      f = c ?? u ?? d;
    if (
      f &&
      !Uw.test(o[f[0].length] ?? "") &&
      !/^\s*[(!]/.test(o.slice(f[0].length))
    ) {
      if (c) i += `${Fw(c[1], c[2], n)}:${Fw(c[3], c[4], n)}`;
      else if (u) i += `${zw(u[1], u[2], t)}:${zw(u[3], u[4], t)}`;
      else if (d) {
        let p = Fw(d[1], d[2], n),
          h = zw(d[3], d[4], t);
        i += p === "#REF!" || h === "#REF!" ? "#REF!" : p + h;
      }
      r += f[0].length;
      continue;
    }
    ((i += a), r++);
  }
  return i;
}
var Xu = Object.freeze({
    rulesPerSheet: 100,
    listItems: 200,
    checks: 1e5,
    customChecks: 5e3,
    displayedIssues: 500,
  }),
  rD = (e) => !!e && typeof e == "object" && !Array.isArray(e),
  Zl = (e, t, n) => {
    if (typeof e != "string" || e.length > t)
      throw Error(`Invalid validation ${n}.`);
    return e;
  };
function eC(e) {
  if (!Array.isArray(e) || e.length > Xu.rulesPerSheet)
    throw Error("At most 100 validation rules are supported.");
  let t = new Set();
  return e.map((n) => {
    if (
      !rD(n) ||
      ![
        "list",
        "number",
        "integer",
        "textLength",
        "date",
        "time",
        "custom",
      ].includes(String(n.kind)) ||
      ![
        "between",
        "notBetween",
        "greater",
        "greaterEqual",
        "less",
        "lessEqual",
        "equal",
        "notEqual",
      ].includes(String(n.operator)) ||
      typeof n.allowBlank != "boolean" ||
      !Array.isArray(n.list) ||
      n.list.length > Xu.listItems ||
      typeof n.min != "number" ||
      !Number.isFinite(n.min) ||
      typeof n.max != "number" ||
      !Number.isFinite(n.max)
    )
      throw Error("Invalid data validation rule.");
    let i = Zl(n.id, 100, "ID");
    if (!i.trim() || t.has(i))
      throw Error("Validation rule IDs must be nonempty and unique.");
    t.add(i);
    let r = Qe(Zl(n.range, 40, "range")),
      o = {
        id: i,
        range: `${Ue(r.firstCol)}${r.firstRow + 1}:${Ue(r.lastCol)}${r.lastRow + 1}`,
        kind: n.kind,
        operator: n.operator,
        min: n.min,
        max: n.max,
        list: n.list.map((a) => Zl(a, 1e3, "list item")),
        allowBlank: n.allowBlank,
        message: Zl(n.message, 500, "error message"),
      };
    if (o.kind === "list" && (!o.list.length || o.list.some((a) => !a.length)))
      throw Error("A list rule needs nonempty values.");
    if (
      !["list", "custom"].includes(o.kind) &&
      ["between", "notBetween"].includes(o.operator) &&
      o.min > o.max
    )
      throw Error("Minimum cannot exceed maximum.");
    if (
      o.kind === "time" &&
      (o.min < 0 || o.min >= 1 || o.max < 0 || o.max >= 1)
    )
      throw Error(
        "Time bounds must be fractions from 0 up to, but excluding, 1.",
      );
    if (
      n.formula !== undefined &&
      ((o.formula = Zl(n.formula, 8192, "formula")),
      !o.formula.startsWith("=") || o.formula.length < 2)
    )
      throw Error("Start the custom validation formula with =.");
    if (o.kind === "custom" && !o.formula)
      throw Error(
        "A custom rule requires a formula relative to the range's top-left cell.",
      );
    if (n.errorStyle !== undefined) {
      if (!["stop", "warning", "information"].includes(String(n.errorStyle)))
        throw Error("Invalid validation error style.");
      o.errorStyle = n.errorStyle;
    }
    for (let a of ["showInputMessage", "showDropdown"])
      if (n[a] !== undefined) {
        if (typeof n[a] != "boolean")
          throw Error("Invalid validation display setting.");
        o[a] = n[a];
      }
    for (let [a, s] of [
      ["inputTitle", 100],
      ["inputMessage", 500],
      ["errorTitle", 100],
    ])
      n[a] !== undefined && (o[a] = Zl(n[a], s, a));
    return o;
  });
}
function tC(e, t, n = Hl) {
  if (t === null || t === "") return e.allowBlank;
  if (e.kind === "custom") return false;
  if (e.kind === "list")
    return e.list.some((r) => r.toLowerCase() === String(t).toLowerCase());
  let i =
    e.kind === "textLength" ? String(t).length : typeof t == "number" ? t : NaN;
  if (
    (["date", "time"].includes(e.kind) &&
      typeof t == "string" &&
      (i = Hi(t, n) ?? NaN),
    !Number.isFinite(i) ||
      (e.kind === "integer" && !Number.isInteger(i)) ||
      (e.kind === "date" &&
        (i < 1 ||
          i > _a.maxSerial + 0.999999999 ||
          (typeof t == "string" && !/[/-]/.test(t)))) ||
      (e.kind === "time" && (i < 0 || i >= 1)))
  )
    return false;
  switch (e.operator) {
    case "between":
      return i >= e.min && i <= e.max;
    case "notBetween":
      return i < e.min || i > e.max;
    case "greater":
      return i > e.min;
    case "greaterEqual":
      return i >= e.min;
    case "less":
      return i < e.min;
    case "lessEqual":
      return i <= e.min;
    case "notEqual":
      return i !== e.min;
    default:
      return i === e.min;
  }
}
function nD(e) {
  let t = Hu(e);
  return (n, i, r) => {
    let o = t(n, i),
      a = false,
      s = "Value does not satisfy the rule.",
      l,
      c = false;
    try {
      if (t.getTypedValue(n, i).type === "error")
        throw Error(`Cell calculation returned ${o}.`);
      if (o === null || o === "") a = r.allowBlank;
      else if (r.kind === "custom") {
        let u = tt(i),
          d = Qe(r.range);
        l = bi(r.formula, u.row - d.firstRow, u.col - d.firstCol);
        let f = t.evaluateFormula(n, i, l);
        if (
          typeof f != "boolean" &&
          (typeof f != "number" || !Number.isFinite(f))
        )
          throw Error(
            `Custom formula must return TRUE/FALSE or a finite number${typeof f == "string" && /^#[A-Z0-9/?!]+$/.test(f) ? `; returned ${f}` : ""}.`,
          );
        ((a = !!f), (s = "Custom formula returned FALSE or zero."));
      } else a = tC(r, o, e.datePolicy);
    } catch (u) {
      ((c = true),
        (s =
          u instanceof Error
            ? u.message
            : "Validation could not be evaluated."));
    }
    if (!(a && !c))
      return {
        sheetId: n,
        sheetName: e.sheets.find((u) => u.id === n).name,
        address: i,
        ruleId: r.id,
        style: c ? "stop" : (r.errorStyle ?? "stop"),
        value: o,
        title: r.errorTitle || "Invalid value",
        message:
          r.message || "The value does not satisfy the data validation rule.",
        reason: s,
        ...(l ? { formula: l } : {}),
        evaluationError: c,
      };
  };
}
var iD = new WeakMap();
function oD(e, t) {
  let n = {
      workbookId: e.id,
      revision: e.revision,
      checks: 0,
      invalidCells: 0,
      violations: 0,
      issues: [],
      complete: true,
      detailsTruncated: false,
      message: "",
    },
    i = nD(e),
    r = new Set(),
    o = 0;
  e: for (let a of e.sheets.filter((s) => !t || s.id === t))
    for (let s of a.validations ?? []) {
      let l = Qe(s.range);
      for (let c = l.firstRow; c <= l.lastRow; c++)
        for (let u = l.firstCol; u <= l.lastCol; u++) {
          if (
            n.checks >= Xu.checks ||
            (s.kind === "custom" && ++o > Xu.customChecks)
          ) {
            n.complete = false;
            break e;
          }
          let d = `${Ue(u)}${c + 1}`;
          n.checks++;
          let f = i(a.id, d, s);
          f &&
            (r.add(`${a.id}:${d}`),
            n.violations++,
            n.issues.length < Xu.displayedIssues && n.issues.push(f));
        }
    }
  return (
    (n.invalidCells = r.size),
    (n.detailsTruncated = n.violations > n.issues.length),
    (n.message = `${n.complete ? "" : "Incomplete inspection: at least "}${n.invalidCells} invalid cells; ${n.violations} rule violations across ${n.checks} checks.${n.complete ? "" : " The 100,000-check or 5,000-custom-formula limit was reached. Narrow validation ranges and inspect again."}${n.detailsTruncated ? " Showing the first 500 violations." : ""}`),
    iD.set(n, { source: JSON.stringify(e), content: JSON.stringify(n) }),
    n
  );
}
function Bw(e) {
  if (!e.sheets.some((n) => n.validations?.length)) return [];
  let t = oD(e);
  return t.invalidCells || !t.complete
    ? [
        `Data validation: ${t.message} Source values were retained; use Data Validation to review.`,
      ]
    : [];
}
var sD = [
    "greater",
    "less",
    "equal",
    "between",
    "contains",
    "duplicate",
    "unique",
    "blank",
    "notBlank",
    "error",
    "notError",
    "formula",
    "colorScale",
    "dataBar",
    "iconSet",
  ],
  aD = (e) => !!e && typeof e == "object" && !Array.isArray(e);
function Gu(e, t, n) {
  if (typeof e != "string" || e.length > t) throw Error(`Invalid ${n}.`);
  return e;
}
function jw(e) {
  if (typeof e != "string" || !/^#[\da-f]{6}$/i.test(e))
    throw Error("Choose a six-digit hexadecimal color.");
  return e;
}
function rC(e) {
  if (!Array.isArray(e) || e.length > 100)
    throw Error("At most 100 conditional formatting rules are supported.");
  let t = new Set(),
    n = new Set();
  return e.map((i) => {
    if (!aD(i) || !sD.includes(i.kind))
      throw Error("Invalid conditional rule.");
    let r = Gu(i.id, 100, "conditional rule ID");
    if (!r.trim() || t.has(r))
      throw Error("Conditional rule IDs must be nonempty and unique.");
    t.add(r);
    let o = Qe(Gu(i.range, 40, "conditional range")),
      a = {
        id: r,
        range: `${Ue(o.firstCol)}${o.firstRow + 1}:${Ue(o.lastCol)}${o.lastRow + 1}`,
        kind: i.kind,
        value: Gu(i.value, 1e3, "conditional value"),
        second: Gu(i.second, 1e3, "conditional second value"),
        fill: jw(i.fill),
        color: jw(i.color),
      };
    if (i.priority !== undefined) {
      if (
        !Number.isInteger(i.priority) ||
        Number(i.priority) < 1 ||
        Number(i.priority) > 100 ||
        n.has(Number(i.priority))
      )
        throw Error(
          "Conditional priorities must be unique integers from 1 to 100.",
        );
      ((a.priority = Number(i.priority)), n.add(a.priority));
    }
    if (i.stopIfTrue !== undefined) {
      if (typeof i.stopIfTrue != "boolean")
        throw Error("Stop if true must be boolean.");
      if (i.stopIfTrue && ["colorScale", "dataBar", "iconSet"].includes(a.kind))
        throw Error(
          "Stop if true is not supported for color scales, data bars, or icon sets.",
        );
      a.stopIfTrue = i.stopIfTrue;
    }
    if (
      (i.formula !== undefined &&
        (a.formula = Gu(i.formula, 4096, "conditional formula")),
      a.kind === "formula" &&
        (!a.formula?.startsWith("=") || !a.formula.slice(1).trim()))
    )
      throw Error("A formula rule requires a formula beginning with =.");
    for (let s of ["minColor", "maxColor"])
      i[s] !== undefined && (a[s] = jw(i[s]));
    if (i.thresholdType !== undefined) {
      if (i.thresholdType !== "percent" && i.thresholdType !== "number")
        throw Error("Invalid icon threshold type.");
      a.thresholdType = i.thresholdType;
    }
    if (i.iconThresholds !== undefined) {
      if (
        !Array.isArray(i.iconThresholds) ||
        i.iconThresholds.length !== 2 ||
        i.iconThresholds.some(
          (s) => typeof s != "number" || !Number.isFinite(s),
        ) ||
        i.iconThresholds[0] >= i.iconThresholds[1] ||
        ((a.thresholdType ?? "percent") === "percent" &&
          (i.iconThresholds[0] < 0 || i.iconThresholds[1] > 100))
      )
        throw Error(
          "Choose two increasing icon thresholds; percentages must be between 0 and 100.",
        );
      a.iconThresholds = [...i.iconThresholds];
    }
    return a;
  });
}
var nC = (e) => !!e && typeof e == "object" && !Array.isArray(e);
function Zu(e) {
  throw Error(e);
}
var lD = (e, t = 200) =>
  typeof e == "string" && e.length <= t
    ? e
    : Zu("Invalid worksheet tool text.");
var cD = (e) => {
    let t = Qe(e);
    return `${Ue(t.firstCol)}${t.firstRow + 1}:${Ue(t.lastCol)}${t.lastRow + 1}`;
  },
  uD = (e, t) => {
    let n = Qe(e),
      i = Qe(t);
    return (
      n.firstRow <= i.lastRow &&
      n.lastRow >= i.firstRow &&
      n.firstCol <= i.lastCol &&
      n.lastCol >= i.firstCol
    );
  };
function Rp(e, t) {
  if (nC(e)) {
    if (e.merges !== undefined) {
      (!Array.isArray(e.merges) || e.merges.length > 1e3) &&
        Zu("At most1,000 merged ranges are supported.");
      let n = [];
      for (let i of e.merges) {
        let r = cD(lD(i, 40)),
          o = Qe(r);
        ((o.lastRow - o.firstRow + 1) * (o.lastCol - o.firstCol + 1) > 500 &&
          Zu("A merged area supports at most 500 cells."),
          n.some((a) => uD(a, r)) && Zu("Merged ranges cannot overlap."),
          n.push(r));
      }
      t.merges = n;
    }
    if (
      (e.validations !== undefined && (t.validations = eC(e.validations)),
      e.conditionalFormats !== undefined &&
        (t.conditionalFormats = rC(e.conditionalFormats)),
      e.protection !== undefined)
    ) {
      let n = e.protection;
      ((!nC(n) ||
        typeof n.enabled != "boolean" ||
        (n.hash !== undefined &&
          (!/^[a-f\d]{64}$/i.test(String(n.hash)) ||
            typeof n.salt != "string" ||
            n.salt.length > 100))) &&
        Zu("Invalid sheet protection."),
        (t.protection = {
          enabled: n.enabled,
          ...(n.hash ? { hash: String(n.hash), salt: String(n.salt) } : {}),
        }));
    }
  }
}
var dD = ["en-US", "en-GB", "de-DE", "fr-FR", "ja-JP", "en-IN"],
  fD = ["USD", "EUR", "GBP", "JPY", "CAD", "AUD", "INR", "CHF"],
  hD = [
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
  pD = {
    date: ["yyyy-mm-dd", "dd/mm/yyyy", "m/d/yyyy", "d mmm yyyy", "d mmmm yyyy"],
    time: ["hh:mm", "hh:mm:ss", "h:mm AM/PM", "hh:mm:ss.000"],
    datetime: [
      "yyyy-mm-dd hh:mm:ss",
      "dd/mm/yyyy hh:mm",
      "m/d/yyyy h:mm AM/PM",
    ],
    duration: ["[h]:mm", "[h]:mm:ss", "[h]:mm:ss.000", "[m]:ss", "[s].000"],
  };
function Lt(e) {
  throw new Error(e);
}
var mD = new Set(["number", "currency", "accounting", "percent"]),
  gD = new Set(["date", "time", "datetime", "duration"]);
function bD(e) {
  return (
    !!e &&
    typeof e == "object" &&
    !Array.isArray(e) &&
    [Object.prototype, null].includes(Object.getPrototypeOf(e))
  );
}
function Vw(e) {
  (!bD(e) ||
    Object.keys(e).some(
      (o) =>
        ![
          "category",
          "locale",
          "currency",
          "decimals",
          "useGrouping",
          "scale",
          "negative",
          "pattern",
        ].includes(o),
    )) &&
    Lt("Number format must be a supported plain settings object.");
  let t = e;
  hD.includes(t.category) || Lt("Choose a supported number format category.");
  let n = t.category,
    i = t.locale ?? "en-US";
  dD.includes(i) || Lt("Choose one of the supported display locales.");
  let r = { category: n, locale: i };
  if (mD.has(n)) {
    if (n === "currency" || n === "accounting") {
      let l = t.currency ?? "USD";
      (fD.includes(l) || Lt("Choose a supported currency."), (r.currency = l));
    } else
      t.currency !== undefined &&
        Lt("Currency applies only to Currency or Accounting formats.");
    let a = t.decimals ?? (r.currency === "JPY" ? 0 : 2);
    ((typeof a != "number" || !Number.isInteger(a) || a < 0 || a > 12) &&
      Lt("Choose 0\u201312 decimal places."),
      (r.decimals = a),
      t.useGrouping !== undefined &&
        typeof t.useGrouping != "boolean" &&
        Lt("The grouping option must be true or false."),
      (r.useGrouping = t.useGrouping ?? n !== "percent"));
    let s = t.scale ?? 1;
    ((![1, 1e3, 1e6].includes(s) || (n === "percent" && s !== 1)) &&
      Lt("Choose units, thousands or millions; percentages use units."),
      (r.scale = s),
      t.negative !== undefined &&
        (typeof t.negative != "string" ||
          !["minus", "parentheses"].includes(t.negative)) &&
        Lt("Choose minus signs or parentheses for negatives."),
      n === "accounting" &&
        t.negative !== undefined &&
        t.negative !== "parentheses" &&
        Lt("Accounting negatives use parentheses."),
      (r.negative =
        n === "accounting" ? "parentheses" : (t.negative ?? "minus")),
      t.pattern !== undefined &&
        Lt("Use the Custom category to enter a format pattern."));
  } else if (
    ([t.currency, t.decimals, t.useGrouping, t.scale, t.negative].some(
      (o) => o !== undefined,
    ) && Lt("These numeric settings do not apply to the chosen category."),
    gD.has(n) || n === "custom")
  ) {
    let o = n === "custom" ? "0.00" : pD[n][0],
      a = t.pattern ?? o;
    ((typeof a != "string" ||
      !a ||
      a.length > 200 ||
      /[\u0000-\u001f\u007f]/.test(a)) &&
      Lt(
        "Enter a format pattern of 1\u2013200 characters without control characters.",
      ),
      (r.pattern = a));
    let s = Ju(r.pattern);
    if (n !== "custom") {
      s.kind !== "temporal" &&
        Lt("Choose a date, time or elapsed-time pattern for this category.");
      let l = s.category;
      l !== n &&
        Lt(
          `This pattern is a ${l} format. Choose the matching category or Custom.`,
        );
    }
  } else
    t.pattern !== undefined && Lt("General does not accept a custom pattern.");
  return r;
}
function yD(e) {
  let t = [""],
    n = false;
  for (let i = 0; i < e.length; i++) {
    let r = e[i];
    if (r === "\\" && i + 1 < e.length) {
      t[t.length - 1] += r + e[++i];
      continue;
    }
    (r === '"' && (n = !n),
      r === ";" && !n ? t.push("") : (t[t.length - 1] += r));
  }
  return n || t.length > 4 ? undefined : t;
}
function vD(e) {
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
    n = (i) => {
      t.hasNumber ? (t.suffix += i) : (t.prefix += i);
    };
  for (let i = 0; i < e.length;) {
    let r = e[i];
    if (r === '"') {
      let o = e.indexOf('"', i + 1);
      if (o < 0) return;
      (n(e.slice(i + 1, o)), (i = o + 1));
    } else if (r === "\\") {
      if (i + 1 >= e.length) return;
      (n(e[i + 1]), (i += 2));
    } else if (r === "[") {
      let o =
        /^\[(?:Black|Blue|Cyan|Green|Magenta|Red|White|Yellow|Color\d+)\]/i.exec(
          e.slice(i),
        );
      if (!o) return;
      i += o[0].length;
    } else if (r === "0" || r === "#") {
      if (t.hasNumber) return;
      let o = /^[0#.,]+/.exec(e.slice(i))[0],
        a = /^(0+|#+0*|#,##0+)(?:\.(0*#*))?$/.exec(o);
      if (!a) return;
      let s = a[2] ?? "";
      if (
        ((t.minDecimals = [...s].filter((l) => l === "0").length),
        (t.maxDecimals = s.length),
        (t.minIntegerDigits = [...a[1]].filter((l) => l === "0").length),
        t.maxDecimals > 20 || t.minIntegerDigits > 21)
      )
        return;
      ((t.grouping = a[1].includes(",")),
        (t.hasNumber = true),
        (i += o.length));
    } else if (r === "%") {
      if (t.percent) return;
      ((t.percent = true), n(r), i++);
    } else if (/[$€£¥ ()+\-]/.test(r)) (n(r), i++);
    else return;
  }
  return t;
}
var Ku = new Map();
function wD(e) {
  let t = e.replace(/"[^"]*"|\\./g, "");
  if (!/[ymdhs]/i.test(t)) return;
  let n = [],
    i = 0;
  for (let l = 0; l < e.length;) {
    let c = e.slice(l),
      u = c[0];
    if (u === '"') {
      let m = e.indexOf('"', l + 1);
      (m < 0 && Lt("The format contains an unclosed quoted literal."),
        n.push({ kind: "literal", text: e.slice(l + 1, m) }),
        (l = m + 1));
      continue;
    }
    if (u === "\\") {
      (c.length < 2 && Lt("A format escape must be followed by a character."),
        n.push({ kind: "literal", text: c[1] }),
        (l += 2));
      continue;
    }
    let d = /^(AM\/PM|A\/P)/i.exec(c);
    if (d) {
      (n.push({ kind: "ampm", text: d[0] }), (l += d[0].length));
      continue;
    }
    let f = /^\[(h{1,2}|m{1,2}|s{1,2})\]/i.exec(c);
    if (f) {
      (n.push({ kind: "elapsed", text: f[1].toLowerCase() }),
        (l += f[0].length));
      continue;
    }
    let p = /^(y+|m+|d+|h+|s+)/i.exec(c);
    if (p) {
      let m = p[0].toLowerCase(),
        g = m[0];
      (((g === "y" && ![2, 4].includes(m.length)) ||
        (g === "m" && m.length > 4) ||
        (["d", "h", "s"].includes(g) && m.length > 2)) &&
        Lt(
          "This date pattern uses unsupported fields. Use yy/yyyy, m\u2013mmmm, d/dd, h/hh or s/ss.",
        ),
        n.push({
          kind: { y: "year", m: "month", d: "day", h: "hour", s: "second" }[g],
          text: m,
        }),
        (l += m.length));
      continue;
    }
    let h = /^\.(0{1,3})(?!0)/.exec(c);
    if (h) {
      ((!n.length ||
        !["second", "elapsed"].includes(n.at(-1).kind) ||
        (n.at(-1).kind === "elapsed" && !n.at(-1).text.startsWith("s")) ||
        i) &&
        Lt(
          "Fractional seconds must follow seconds and use 1\u20133 zero digits.",
        ),
        (i = h[1].length),
        n.push({ kind: "fraction", text: h[1] }),
        (l += h[0].length));
      continue;
    }
    if (/^[ /:,.-]$/.test(u)) {
      (n.push({ kind: "literal", text: u }), l++);
      continue;
    }
    Lt(
      "Unsupported date/time code. Quote literal words and use one supported date/time pattern.",
    );
  }
  let r = n.filter((l) => l.kind !== "literal");
  r.forEach((l, c) => {
    l.kind === "month" &&
      l.text.length <= 2 &&
      (r[c - 1]?.kind === "hour" ||
        (r[c - 1]?.kind === "elapsed" && r[c - 1].text.startsWith("h")) ||
        r[c + 1]?.kind === "second") &&
      (l.kind = "minute");
  });
  let o = r.some((l) => ["year", "month", "day"].includes(l.kind)),
    a = r.some((l) => ["hour", "minute", "second"].includes(l.kind)),
    s = r.filter((l) => l.kind === "elapsed");
  if (s.length) {
    (s.length !== 1 ||
      o ||
      r.some((c) => ["hour", "ampm"].includes(c.kind)) ||
      r[0].kind !== "elapsed") &&
      Lt(
        "Elapsed-time patterns start with one [h], [m] or [s] field and cannot include calendar or AM/PM fields.",
      );
    let l = s[0].text[0];
    ((l === "s" && a) || (l === "m" && r.some((c) => c.kind === "minute"))) &&
      Lt("Use [h]:mm:ss, [m]:ss or [s] for elapsed time.");
  }
  return (
    r.some((l) => l.kind === "ampm") &&
      !r.some((l) => l.kind === "hour") &&
      Lt("AM/PM requires an hour field."),
    !o &&
      !a &&
      !s.length &&
      Lt("The format does not contain a date or time field."),
    {
      kind: "temporal",
      category: s.length
        ? "duration"
        : o && a
          ? "datetime"
          : o
            ? "date"
            : "time",
      tokens: n,
      fraction: i,
    }
  );
}
function xD(e) {
  let t = "",
    n = 1,
    i = false;
  for (let o = 0; o < e.length;) {
    let a = e.slice(o);
    if (a[0] === '"') {
      let s = e.indexOf('"', o + 1);
      (s < 0 && Lt("The format contains an unclosed quoted literal."),
        (t += e.slice(o, s + 1)),
        (o = s + 1));
    } else if (a[0] === "\\")
      (a.length < 2 && Lt("A format escape must be followed by a character."),
        (t += a.slice(0, 2)),
        (o += 2));
    else if (/^[0#]/.test(a)) {
      i && Lt("Only one numeric field per custom section is supported.");
      let s = /^[0#.,]+/.exec(a)[0],
        l = /,+$/.exec(s)?.[0] ?? "";
      (l.length > 2 &&
        Lt("Custom display scaling supports at most two trailing commas."),
        (n = 1e3 ** l.length),
        (t += s.slice(0, s.length - l.length)),
        (o += s.length),
        (i = true));
    } else ((t += a[0]), o++);
  }
  /\[/.test(t.replace(/"[^"]*"|\\./g, "")) &&
    Lt(
      "Custom colors, conditions and locale tags are not supported. Use conditional formatting or the locale selector.",
    );
  let r = vD(t);
  return (
    r ||
      Lt(
        "Unsupported custom format. Use 0/# digits, grouping, decimals, percent, quoted literals and positive/negative/zero sections.",
      ),
    r.maxDecimals > 12 &&
      Lt("Custom formats support at most 12 decimal places."),
    { parsed: r, divisor: n }
  );
}
function Ju(e) {
  let t = Ku.get(e);
  if (t) return t;
  (typeof e != "string" ||
    !e.length ||
    e.length > 200 ||
    /[\u0000-\u001f\u007f]/.test(e)) &&
    Lt("Enter a supported format pattern of 1\u2013200 characters.");
  let n,
    i = yD(e);
  i || Lt("A format has at most four sections and must close quoted literals.");
  let r = i.length === 1 ? wD(e) : undefined;
  return (
    r
      ? (n = r)
      : (i.length === 4 &&
          !kD(i[3]) &&
          Lt(
            "The text section supports @ and quoted or escaped literals only.",
          ),
        (n = {
          kind: "numeric",
          sections: i.slice(0, 3).map(xD),
          ...(i.length === 4 ? { textSection: i[3] } : {}),
        })),
    Ku.size >= 256 && Ku.delete(Ku.keys().next().value),
    Ku.set(e, n),
    n
  );
}
function kD(e) {
  return e.replace(/"[^"]*"|\\./g, "").replace(/[@ ]/g, "") === "";
}
function _D(e, t) {
  return e.replace(
    /"([^"]*)"|\\(.)|@/g,
    (n, i, r) => i ?? r ?? (n === "@" ? t : n),
  );
}
var iC = (e) =>
  new Intl.NumberFormat(e).formatToParts(1.1).find((t) => t.type === "decimal")
    ?.value ?? ".";
function ED(e, t, n) {
  let i = e < 0 && t.sections.length > 1,
    { parsed, divisor } =
      t.sections[e === 0 && t.sections.length > 2 ? 2 : i ? 1 : 0];
  if (!parsed.hasNumber) return parsed.prefix + parsed.suffix;
  let a = (Math.abs(e) / divisor) * (parsed.percent ? 100 : 1);
  if (!Number.isFinite(a)) return "#NUM!";
  let s = new Intl.NumberFormat(n, {
    minimumFractionDigits: parsed.minDecimals,
    maximumFractionDigits: parsed.maxDecimals,
    minimumIntegerDigits: Math.max(1, parsed.minIntegerDigits),
    useGrouping: parsed.grouping,
  }).format(a);
  return (
    parsed.minIntegerDigits ||
      (s === "0" && !parsed.minDecimals
        ? (s = "")
        : s.startsWith(`0${iC(n)}`) && (s = s.slice(1))),
    (e < 0 && !i ? "-" : "") + parsed.prefix + s + parsed.suffix
  );
}
function SD(e, t, n, i) {
  let r = t.category === "duration",
    o = 2958465 - (i === "1904" && !r ? 1462 : 0);
  if (Math.abs(e) >= o + 1 || (!r && e < 0)) return "#NUM!";
  let a = t.category !== "date",
    s = 10 ** t.fraction,
    l = t.tokens.some(
      (P) =>
        P.kind === "second" || (P.kind === "elapsed" && P.text.startsWith("s")),
    ),
    c = a
      ? l
        ? Math.round(Math.abs(e) * 86400 * s) / s
        : Math.floor(Math.abs(e) * 86400)
      : Math.floor(e) * 86400,
    u = Math.floor(c / 86400),
    d = c - u * 86400;
  if (u > o) return "#NUM!";
  let f = Math.round(d * 1e3),
    p = Math.floor(f / 36e5),
    h = Math.floor(f / 6e4) % 60,
    m = Math.floor(f / 1e3) % 60,
    g = Math.floor((f % 1e3) / 10 ** (3 - t.fraction)),
    b = i === "1900" && u === 60,
    w =
      i === "1904"
        ? Date.UTC(1904, 0, 1) + u * 864e5
        : Date.UTC(1899, 11, 31) + (u > 60 ? u - 1 : u) * 864e5,
    _ = new Date(w),
    S = b ? 1900 : _.getUTCFullYear(),
    O = b ? 1 : _.getUTCMonth(),
    E = b ? 29 : _.getUTCDate(),
    R = t.tokens.some((P) => P.kind === "ampm"),
    I = (P, T) => String(P).padStart(T, "0"),
    U = t.tokens
      .map((P) => {
        switch (P.kind) {
          case "literal":
            return P.text;
          case "year":
            return P.text.length === 2 ? I(S % 100, 2) : I(S, 4);
          case "month":
            return P.text.length < 3
              ? I(O + 1, P.text.length)
              : new Intl.DateTimeFormat(n, {
                  month: P.text.length === 3 ? "short" : "long",
                  timeZone: "UTC",
                }).format(new Date(Date.UTC(2e3, O, 1)));
          case "day":
            return I(E, P.text.length);
          case "hour":
            return I(R ? p % 12 || 12 : p, P.text.length);
          case "minute":
            return I(h, P.text.length);
          case "second":
            return I(m, P.text.length);
          case "fraction":
            return iC(n) + I(g, t.fraction);
          case "ampm": {
            let T = p < 12 ? "AM" : "PM",
              A = P.text.length === 3 ? T[0] : T;
            return P.text === P.text.toLowerCase() ? A.toLowerCase() : A;
          }
          case "elapsed":
            return I(
              Math.floor(
                c / (P.text[0] === "h" ? 3600 : P.text[0] === "m" ? 60 : 1),
              ),
              P.text.length,
            );
        }
      })
      .join("");
  return (r && e < 0 ? "-" : "") + U;
}
function oC(e, t, n = "1900") {
  let i = Vw(t);
  if (e === null) return { text: "" };
  if (typeof e == "boolean") return { text: e ? "TRUE" : "FALSE" };
  if (typeof e == "string") {
    if (i.category === "custom" && !/^#[A-Z0-9/?!]+$/.test(e)) {
      let c = Ju(i.pattern);
      if (c.kind === "numeric" && c.textSection !== undefined)
        return { text: _D(c.textSection, e) };
    }
    return { text: e };
  }
  if (!Number.isFinite(e)) return { text: "#NUM!" };
  if (i.category === "general")
    return {
      text: new Intl.NumberFormat(i.locale, {
        maximumSignificantDigits: 15,
        useGrouping: false,
      }).format(e === 0 ? 0 : e),
    };
  if (i.pattern) {
    let c = Ju(i.pattern);
    return {
      text: c.kind === "temporal" ? SD(e, c, i.locale, n) : ED(e, c, i.locale),
    };
  }
  let r = e / (i.scale ?? 1),
    o = i.scale === 1e3 ? " K" : i.scale === 1e6 ? " M" : "";
  if (i.category === "percent" && !Number.isFinite(r * 100))
    return { text: "#NUM!" };
  let a = {
    minimumFractionDigits: i.decimals,
    maximumFractionDigits: i.decimals,
    useGrouping: i.useGrouping,
  };
  if (i.category === "accounting") {
    let u = new Intl.NumberFormat(i.locale, {
        ...a,
        style: "currency",
        currency: i.currency,
      }).formatToParts(Math.abs(r)),
      d = u.find((m) => m.type === "currency")?.value ?? i.currency,
      f =
        u.findIndex((m) => m.type === "currency") <
        u.findIndex((m) => m.type === "integer")
          ? "start"
          : "end",
      p = new Intl.NumberFormat(i.locale, a).format(Math.abs(r)),
      h = (r === 0 ? "-" : r < 0 ? `(${p})` : p) + o;
    return {
      text: f === "start" ? `${d} ${h}` : `${h} ${d}`,
      accounting: { symbol: d, amount: h, symbolPosition: f },
    };
  }
  let s = r < 0 && i.negative === "parentheses",
    l = new Intl.NumberFormat(i.locale, {
      ...a,
      ...(i.category === "currency"
        ? { style: "currency", currency: i.currency }
        : i.category === "percent"
          ? { style: "percent" }
          : {}),
    }).format(s ? Math.abs(r) : r);
  return { text: (s ? `(${l})` : l) + o };
}
var K2 = Bi(LI(), 1);
function UI(e) {
  let [t, n = "0"] = Math.abs(e).toString().split("e"),
    [i, r = ""] = t.split(".");
  return {
    coefficient: BigInt(i + r) * (e < 0 ? -1n : 1n),
    exponent: Number(n) - r.length,
  };
}
function FI(e) {
  let t = new Map();
  for (let o of e) {
    if (!Number.isFinite(o)) return;
    let { coefficient, exponent } = UI(o);
    t.set(exponent, (t.get(exponent) ?? 0n) + coefficient);
  }
  let n = [...t].filter(([, o]) => o !== 0n);
  if (!n.length) return { coefficient: 0n, exponent: 0 };
  let i = n.reduce((o, [a]) => Math.min(o, a), 1 / 0);
  return {
    coefficient: n.reduce((o, [a, s]) => o + s * 10n ** BigInt(a - i), 0n),
    exponent: i,
  };
}
function oi(e) {
  let t = FI(e);
  return t ? +`${t.coefficient}e${t.exponent}` : NaN;
}
function sn(e) {
  if (!e.length) return NaN;
  let t = FI(e);
  if (!t) return NaN;
  let n = t.coefficient < 0n ? "-" : "",
    i = t.coefficient < 0n ? -t.coefficient : t.coefficient,
    r = BigInt(e.length),
    o = i / r,
    a = (((i % r) * 10n ** 40n) / r).toString().padStart(40, "0");
  return +`${n}${o}.${a}e${t.exponent}`;
}
function Nm(e) {
  if (e.some((i) => !Number.isFinite(i))) return NaN;
  if (e.some((i) => i === 0)) return 0;
  let t = 1n,
    n = 0;
  for (let i of e) {
    let r = UI(i);
    if (((t *= r.coefficient), (n += r.exponent), e.length > 255)) {
      let o = t < 0n,
        a = (o ? -t : t).toString();
      a.length > 80 &&
        ((n += a.length - 80), (t = BigInt(a.slice(0, 80)) * (o ? -1n : 1n)));
    }
  }
  return +`${t}e${n}`;
}
var u2 = Symbol("not-arithmetic"),
  tj = new Set([
    "ROUND",
    "ROUNDUP",
    "ROUNDDOWN",
    "TRUNC",
    "MROUND",
    "CEILING",
    "FLOOR",
    "SUMPRODUCT",
  ]);
function c2(e) {
  let [t, n = "0"] = Math.abs(e).toString().split("e"),
    [i, r = ""] = t.split(".");
  return { coefficient: BigInt(i + r), exponent: Number(n) - r.length };
}
function zI(e, t, n) {
  let i = e / t,
    r = e % t;
  return (
    i +
    (r !== 0n && (n === "away" || (n === "nearest" && r * 2n >= t)) ? 1n : 0n)
  );
}
function rj(e, t, n) {
  if (((t = Math.trunc(t)), e === 0 || t >= 324)) return e;
  if (t <= -309) return n === "away" ? 1 / 0 : 0;
  let i = c2(e),
    r = i.exponent + t;
  if (r >= 0) return e;
  let o = zI(i.coefficient, 10n ** BigInt(-r), n);
  return Math.sign(e) * +`${o}e${-t}`;
}
function nj(e, t, n) {
  let i = c2(e),
    r = c2(t),
    o = Math.min(i.exponent, r.exponent),
    a = i.coefficient * 10n ** BigInt(i.exponent - o),
    s = r.coefficient * 10n ** BigInt(r.exponent - o),
    l = zI(a, s, n);
  return Math.sign(e) * +`${l * r.coefficient}e${r.exponent}`;
}
function BI(e, t, n) {
  if (!tj.has(e)) return u2;
  let { retrieve, issue, isIssue } = n,
    a = t.map(retrieve),
    s = (f) => (Number.isFinite(f) ? (f === 0 ? 0 : f) : issue("#NUM!"));
  if (e === "SUMPRODUCT") {
    if (!t.length || t.length > 255) return issue("#VALUE!");
    let f = a.map((g) =>
        Array.isArray(g) ? (g.every(Array.isArray) ? g : [g]) : [[g]],
      ),
      p = f[0].length,
      h = f[0][0]?.length ?? 0;
    if (
      !p ||
      !h ||
      f.some((g) => g.length !== p || g.some((b) => b.length !== h))
    )
      return issue("#VALUE!");
    if (p * h * f.length > 1e5) return issue("#LIMIT!");
    let m = [];
    for (let g = 0; g < p; g++)
      for (let b = 0; b < h; b++) {
        let w = [];
        for (let S of f) {
          let O = S[g][b];
          if (isIssue(O)) return O;
          w.push(typeof O == "number" ? O : 0);
        }
        let _ = Nm(w);
        if (!Number.isFinite(_)) return issue("#NUM!");
        m.push(_);
      }
    return s(oi(m));
  }
  if (t.length < (e === "TRUNC" ? 1 : 2) || t.length > 2)
    return issue("#VALUE!");
  let l = [];
  for (let f of a) {
    if (isIssue(f)) return f;
    if (
      Array.isArray(f) ||
      (typeof f == "string" &&
        !/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?$/i.test(f.trim()))
    )
      return issue("#VALUE!");
    let p = f == null ? 0 : Number(f);
    if (!Number.isFinite(p)) return issue("#NUM!");
    l.push(p);
  }
  let [c, u = 0] = l;
  if (["ROUND", "ROUNDUP", "ROUNDDOWN", "TRUNC"].includes(e))
    return s(
      rj(c, u, e === "ROUND" ? "nearest" : e === "ROUNDUP" ? "away" : "toward"),
    );
  if (c === 0 || u === 0) return 0;
  if (
    (e === "MROUND" && Math.sign(c) !== Math.sign(u)) ||
    (e === "FLOOR" && c > 0 && u < 0)
  )
    return issue("#NUM!");
  let d =
    e === "MROUND"
      ? "nearest"
      : e === "CEILING"
        ? c < 0 && u > 0
          ? "toward"
          : "away"
        : c < 0 && u > 0
          ? "away"
          : "toward";
  return s(nj(c, u, d));
}
var d2 = ["LINEST", "LOGEST", "TREND", "GROWTH"],
  f2 = Symbol("not-regression"),
  wd = {
    values: 4e4,
    observations: 1e4,
    predictors: 32,
    work: 4e6,
    output: 1e4,
  },
  Li = (e, t) => {
    let n = 0,
      i = 0;
    for (let r = 0; r < e.length; r++) {
      let o = e[r] * t[r],
        a = n + o;
      ((i += Math.abs(n) >= Math.abs(o) ? n - a + o : o - a + n), (n = a));
    }
    return n + i;
  };
function jI(e, t, n) {
  if (!d2.includes(e)) return f2;
  let i = (Y = "#VALUE!") => {
      throw n.issue(Y);
    },
    r = (Y) => (Number.isFinite(Y) ? Y : i("#NUM!"));
  (!t.length || t.length > 4 || t[0] == null) && i();
  let o = 0,
    a = (Y) => {
      if (n.isIssue(Y)) throw Y;
      let ne = Array.isArray(Y) ? (Y.every(Array.isArray) ? Y : [Y]) : [[Y]],
        ae = ne[0]?.length ?? 0;
      return (
        (!ne.length || !ae || ne.some((le) => le.length !== ae)) && i(),
        (o += ne.length * ae),
        o > wd.values && i("#LIMIT!"),
        ne.map((le) =>
          le.map((de) => {
            if (n.isIssue(de)) throw de;
            return (typeof de != "number" && i(), r(de));
          }),
        )
      );
    },
    s = (Y, ne) => {
      if (Y == null) return ne;
      let ae = n.retrieve(Y);
      if (n.isIssue(ae)) throw ae;
      return ae == null
        ? false
        : typeof ae == "boolean"
          ? ae
          : typeof ae == "number"
            ? r(ae) !== 0
            : typeof ae == "string" && /^(true|false)$/i.test(ae)
              ? /^true$/i.test(ae)
              : i();
    },
    l = e === "TREND" || e === "GROWTH",
    c = e === "LOGEST" || e === "GROWTH",
    u = s(t[l ? 3 : 2], true),
    d = !l && s(t[3], false),
    f = a(n.retrieve(t[0])),
    p = f.length * f[0].length;
  p > wd.observations && i("#LIMIT!");
  let h = f.flat();
  c && h.some((Y) => Y <= 0) && i("#NUM!");
  let m = c ? h.map(Math.log) : h,
    g =
      t[1] == null
        ? f.map((Y, ne) => Y.map((ae, le) => ne * Y.length + le + 1))
        : a(n.retrieve(t[1])),
    b = g.length === f.length && g[0].length === f[0].length,
    w = f.length === 1,
    _;
  if (b) _ = g.flat().map((Y) => [Y]);
  else if (f[0].length === 1 && g.length === p) _ = g;
  else if (w && g[0].length === p)
    _ = Array.from({ length: p }, (Y, ne) => g.map((ae) => ae[ne]));
  else return i("#REF!");
  let S = _[0].length;
  (S > wd.predictors || p * S * S > wd.work) && i("#LIMIT!");
  let O = Array.from({ length: S }, (Y, ne) =>
      u ? sn(_.map((ae) => ae[ne])) : 0,
    ),
    E = u ? sn(m) : 0,
    R = m.map((Y) => r(Y - E)),
    I = Math.max(...R.map(Math.abs)) || 1,
    U = R.map((Y) => Y / I),
    P = [],
    T = Array.from({ length: S }, (Y, ne) => {
      let ae = _.map((de) => r(de[ne] - O[ne])),
        le = Math.max(...ae.map(Math.abs)) || 1;
      return (P.push(le), ae.map((de) => de / le));
    }),
    A = T.map((Y) => [...Y]),
    L = Array.from({ length: S }, (Y, ne) => ne),
    y = Array.from({ length: S }, () => Array(S).fill(0)),
    v = [],
    k = [],
    x = 1e-12 * Math.max(1, ...T.map((Y) => Math.sqrt(Li(Y, Y)))),
    C = 0;
  for (let Y = 0; Y < S; Y++) {
    let ne = Y;
    for (let le = Y + 1; le < S; le++)
      Li(T[le], T[le]) > Li(T[ne], T[ne]) && (ne = le);
    (([T[Y], T[ne]] = [T[ne], T[Y]]), ([L[Y], L[ne]] = [L[ne], L[Y]]));
    for (let le = 0; le < Y; le++)
      [y[le][Y], y[le][ne]] = [y[le][ne], y[le][Y]];
    let ae = Math.sqrt(Li(T[Y], T[Y]));
    if (ae <= x) break;
    ((y[Y][Y] = ae),
      v.push(T[Y].map((le) => le / ae)),
      k.push(Li(v[Y], U)),
      C++);
    for (let le = Y + 1; le < S; le++)
      for (let de = 0; de < 2; de++) {
        let Se = Li(v[Y], T[le]);
        ((y[Y][le] += Se), (T[le] = T[le].map((ge, _e) => ge - Se * v[Y][_e])));
      }
  }
  let N = (Y) => {
      let ne = Array(C).fill(0);
      for (let ae = C - 1; ae >= 0; ae--) {
        let le = Y[ae];
        for (let de = ae + 1; de < C; de++) le -= y[ae][de] * ne[de];
        ne[ae] = r(le / y[ae][ae]);
      }
      return ne;
    },
    F = Array(S).fill(0);
  N(k).forEach((Y, ne) => (F[L[ne]] = Y));
  let z = F.map((Y, ne) => r(Y * (I / P[ne]))),
    j = u ? r(E - Li(z, O)) : 0,
    q = (Y) => {
      let ne = r(Math.exp(Y));
      return ne > 0 ? ne : i("#NUM!");
    };
  if (l) {
    let Y = t[2] == null ? g : a(n.retrieve(t[2])),
      ne,
      ae,
      le;
    if (S === 1)
      ((ne = Y.flat().map((Se) => [Se])), (ae = Y.length), (le = Y[0].length));
    else if (!w && Y[0].length === S) ((ne = Y), (ae = ne.length), (le = 1));
    else if (w && Y.length === S)
      ((ne = Array.from({ length: Y[0].length }, (Se, ge) =>
        Y.map((_e) => _e[ge]),
      )),
        (ae = 1),
        (le = ne.length));
    else return i("#REF!");
    ae * le > wd.output && i("#LIMIT!");
    let de = ne.map((Se) => {
      let ge = r(
        E +
          Li(
            F,
            Se.map((_e, Ge) => r(_e - O[Ge]) / P[Ge]),
          ) *
            I,
      );
      return c ? q(ge) : ge;
    });
    return Array.from({ length: ae }, (Se, ge) =>
      de.slice(ge * le, (ge + 1) * le),
    );
  }
  let D = [...z]
    .reverse()
    .concat(j)
    .map((Y) => (c ? q(Y) : Y));
  if (!d) return [D];
  let W = U.map(
      (Y, ne) =>
        Y -
        Li(
          F,
          A.map((ae) => ae[ne]),
        ),
    ),
    $ = Li(W, W),
    M = Li(U, U);
  $ <= 1e-26 * Math.max(M, Number.MIN_VALUE) && ($ = 0);
  let B = p - C - Number(u),
    V = B > 0 ? Math.sqrt($ / B) : null,
    G = () => n.issue("#N/A"),
    H = () => n.issue("#DIV/0!"),
    K = Array.from({ length: C }, (Y, ne) =>
      N(Array.from({ length: C }, (ae, le) => (le === ne ? 1 : 0))),
    ),
    te = Array(S).fill(0);
  for (let Y = 0; Y < C; Y++)
    te[L[Y]] =
      V === null
        ? H()
        : r(Math.hypot(...K.map((ne) => ne[Y])) * V * (I / P[L[Y]]));
  let Q = L.slice(0, C).map((Y) => O[Y] / P[Y]),
    Z = 1 / p + K.reduce((Y, ne) => Y + Li(Q, ne) ** 2, 0),
    J = [...te]
      .reverse()
      .concat(u ? (V === null ? H() : r(Math.sqrt(Z) * V * I)) : G()),
    ee = Math.max(0, M - $),
    re = (Y) => r(Y * I * I),
    oe = (Y, ne) => [Y, ne, ...Array.from({ length: S - 1 }, G)];
  return [
    D,
    J,
    oe(M ? Math.max(0, Math.min(1, ee / M)) : H(), V === null ? H() : r(V * I)),
    oe(B > 0 && C && $ ? r(ee / C / ($ / B)) : H(), B),
    oe(re(ee), re($)),
  ];
}
var h2 = ["T.TEST", "Z.TEST"],
  p2 = Symbol("not-inference");
function WI(e, t, n) {
  if (!h2.includes(e)) return p2;
  let i = (I = "#VALUE!") => {
      throw n.issue(I);
    },
    r = (I) => (Number.isFinite(I) ? I : i("#NUM!"));
  ((e === "T.TEST" && t.length !== 4) ||
    (e === "Z.TEST" && (t.length < 2 || t.length > 3))) &&
    i();
  let o = t.map(n.retrieve),
    a = 0,
    s = (I) => {
      if (Array.isArray(I)) return I.flatMap(s);
      if ((++a > 4e4 && i("#LIMIT!"), n.isIssue(I))) throw I;
      return (typeof I == "number" && r(I), [I]);
    },
    l = (I) => {
      if (n.isIssue(I)) throw I;
      if (Array.isArray(I) || I == null || (typeof I == "string" && !I.trim()))
        return i();
      let U = Number(I);
      return Number.isFinite(U) ? U : i();
    },
    c = (I) => I.filter((U) => typeof U == "number"),
    u = (I) => {
      if (I.length < 2) return i("#DIV/0!");
      let U = sn(I);
      return {
        mean: U,
        variance: r(oi(I.map((P) => (P - U) ** 2)) / (I.length - 1)),
      };
    },
    d = (I) => {
      if (n.isIssue(I)) throw I;
      return typeof I != "number" ||
        !Number.isFinite(I) ||
        I < -1e-12 ||
        I > 1 + 1e-12
        ? i("#NUM!")
        : Math.max(0, Math.min(1, I));
    },
    f = s(o[0]);
  if (e === "Z.TEST") {
    let I = c(f),
      U = I.length;
    if (!U) return i("#N/A");
    let P = l(o[1]),
      T = Math.max(Math.abs(P), ...I.map(Math.abs)) || 1,
      A = I.map((k) => k / T),
      L = t[2] == null ? Math.sqrt(u(A).variance) : l(o[2]) / T;
    if (L < 0) return i("#NUM!");
    if (!L) return i("#DIV/0!");
    let y = r(((sn(A) - P / T) / L) * Math.sqrt(U)),
      v = n.distribution("ERFC", [y / Math.SQRT2]);
    if (n.isIssue(v)) throw v;
    return typeof v == "number" ? d(v / 2) : i("#NUM!");
  }
  let p = s(o[1]),
    h = Math.trunc(l(o[2])),
    m = Math.trunc(l(o[3]));
  ((![1, 2].includes(h) || ![1, 2, 3].includes(m)) && i("#NUM!"),
    m === 1 && f.length !== p.length && i("#N/A"));
  let g, b;
  if (m === 1) {
    let I = f
      .map((U, P) => [U, p[P]])
      .filter((U) => U.every((P) => typeof P == "number"));
    ((g = I.map((U) => U[0])), (b = I.map((U) => U[1])));
  } else ((g = c(f)), (b = c(p)));
  (g.length < 2 || b.length < 2) && i("#DIV/0!");
  let w = Math.max(...g.map(Math.abs), ...b.map(Math.abs)) || 1;
  ((g = g.map((I) => I / w)), (b = b.map((I) => I / w)));
  let _, S, O;
  if (m === 1) {
    let I = u(g.map((U, P) => U - b[P]));
    ((_ = I.mean), (S = I.variance / g.length), (O = g.length - 1));
  } else {
    let I = u(g),
      U = u(b);
    if (((_ = I.mean - U.mean), m === 2))
      ((O = g.length + b.length - 2),
        (S =
          (((g.length - 1) * I.variance + (b.length - 1) * U.variance) / O) *
          (1 / g.length + 1 / b.length)));
    else {
      let P = I.variance / g.length,
        T = U.variance / b.length;
      ((S = P + T),
        (O = S ** 2 / (P ** 2 / (g.length - 1) + T ** 2 / (b.length - 1))));
    }
  }
  S || i("#DIV/0!");
  let E = r(Math.abs(_) / Math.sqrt(S));
  r(O);
  let R = n.distribution("BETA.DIST", [O / (O + E * E), O / 2, 0.5, true]);
  return (d(R) * h) / 2;
}
var m2 = ["CONVERT", "SLOPE", "RSQ", "STEYX", "PEARSON"],
  g2 = Symbol("not-specialist"),
  xd = [],
  Xe = (e, t, n, i = false, r = 0, o = 1) => {
    for (let a of e.split("|"))
      xd.push({
        symbol: a,
        dimension: t,
        factor: n,
        metric: i,
        offset: r,
        power: o,
      });
  };
Xe("g", "mass", 0.001, true);
Xe("lbm", "mass", 0.45359237);
Xe("ozm", "mass", 0.45359237 / 16);
Xe("grain", "mass", 0.45359237 / 7e3);
Xe("stone", "mass", 0.45359237 * 14);
Xe("ton", "mass", 0.45359237 * 2e3);
Xe("uk_ton|LTON|brton", "mass", 0.45359237 * 2240);
Xe("cwt|shweight", "mass", 0.45359237 * 100);
Xe("uk_cwt|lcwt|hweight", "mass", 0.45359237 * 112);
Xe("m", "length", 1, true);
Xe("in", "length", 0.0254);
Xe("ft", "length", 0.3048);
Xe("yd", "length", 0.9144);
Xe("mi", "length", 1609.344);
Xe("Nmi", "length", 1852);
Xe("ang", "length", 1e-10, true);
Xe("Picapt|Pica", "length", 0.0254 / 72);
Xe("pica", "length", 0.0254 / 6);
Xe("sec|s", "time", 1, true);
Xe("mn|min", "time", 60);
Xe("hr", "time", 3600);
Xe("day|d", "time", 86400);
Xe("yr", "time", 365.25 * 86400);
Xe("Pa|p", "pressure", 1, true);
Xe("atm|at", "pressure", 101325);
Xe("psi", "pressure", 6894.757293168361);
Xe("N", "force", 1, true);
Xe("dyn|dy", "force", 1e-5, true);
Xe("lbf", "force", 4.4482216152605);
Xe("J", "energy", 1, true);
Xe("e", "energy", 1e-7, true);
Xe("c", "energy", 4.184, true);
Xe("cal", "energy", 4.1868, true);
Xe("Wh|wh", "energy", 3600, true);
Xe("BTU|btu", "energy", 1055.05585262);
Xe("W|w", "power", 1, true);
Xe("HP|h", "power", 745.6998715822702);
Xe("PS", "power", 735.49875);
Xe("T", "magnetism", 1, true);
Xe("ga", "magnetism", 1e-4, true);
Xe("K|kel", "temperature", 1, true);
Xe("C|cel", "temperature", 1, false, 273.15);
Xe("F|fah", "temperature", 5 / 9, false, 273.15 - 160 / 9);
Xe("Rank", "temperature", 5 / 9);
Xe("Reau", "temperature", 1.25, false, 273.15);
Xe("l|L|lt", "volume", 0.001, true);
Xe("gal", "volume", 0.003785411784);
Xe("qt", "volume", 0.003785411784 / 4);
Xe("pt|us_pt", "volume", 0.003785411784 / 8);
Xe("cup", "volume", 0.003785411784 / 16);
Xe("oz", "volume", 0.003785411784 / 128);
Xe("tbs", "volume", 0.003785411784 / 256);
Xe("tsp", "volume", 0.003785411784 / 768);
Xe("tspm", "volume", 5e-6);
Xe("uk_gal", "volume", 0.00454609);
Xe("uk_qt", "volume", 0.00454609 / 4);
Xe("uk_pt", "volume", 0.00454609 / 8);
Xe("barrel", "volume", 0.003785411784 * 42);
for (let [e, t] of [
  ["m", 1],
  ["in", 0.0254],
  ["ft", 0.3048],
  ["yd", 0.9144],
  ["mi", 1609.344],
  ["Nmi", 1852],
  ["ang", 1e-10],
])
  (Xe(`${e}2|${e}^2`, "area", t ** 2, e === "m" || e === "ang", 0, 2),
    Xe(`${e}3|${e}^3`, "volume", t ** 3, e === "m" || e === "ang", 0, 3));
Xe("ar", "area", 100, true);
Xe("ha", "area", 1e4);
Xe("uk_acre", "area", 4046.8564224);
Xe("us_acre", "area", (1200 / 3937) ** 2 * 43560);
Xe("bit", "information", 1, true);
Xe("byte", "information", 8, true);
Xe("m/s|m/sec", "speed", 1, true);
Xe("m/h|m/hr", "speed", 1 / 3600, true);
Xe("mph", "speed", 1609.344 / 3600);
Xe("kn", "speed", 1852 / 3600);
var bQ = Object.freeze(xd.map((e) => Object.freeze(e))),
  VI = {
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
  ij = { ki: 10, Mi: 20, Gi: 30, Ti: 40, Pi: 50, Ei: 60, Zi: 70, Yi: 80 };
function qI(e) {
  let t = xd.find((n) => n.symbol === e);
  if (t) return { ...t };
  for (let [n, i] of Object.entries(ij)) {
    if (!e.startsWith(n)) continue;
    let r = xd.find(
      (o) => o.symbol === e.slice(n.length) && o.dimension === "information",
    );
    if (r) return { ...r, symbol: e, factor: r.factor * 2 ** i };
  }
  for (let n of Object.keys(VI).sort((i, r) => r.length - i.length)) {
    if (!e.startsWith(n)) continue;
    let i = xd.find((r) => r.symbol === e.slice(n.length) && r.metric);
    if (i)
      return { ...i, symbol: e, factor: i.factor * 10 ** (VI[n] * i.power) };
  }
}
function XI(e, t, n) {
  if (!m2.includes(e)) return g2;
  let i = (P = "#VALUE!") => {
    throw n.issue(P);
  };
  (t.length !== (e === "CONVERT" ? 3 : 2) || t.some((P) => P == null)) && i();
  let r = t.map(n.retrieve),
    o = 0,
    a = (P) => {
      if (Array.isArray(P)) {
        for (let T of P) a(T);
        return;
      }
      if ((++o > 4e4 && i("#LIMIT!"), n.isIssue(P))) throw P;
      typeof P == "number" && !Number.isFinite(P) && i("#NUM!");
    };
  r.forEach(a);
  let s = (P) => (Number.isFinite(P) ? P : i("#NUM!"));
  if (e === "CONVERT") {
    let [P, T, A] = r;
    (typeof T != "string" ||
      typeof A != "string" ||
      Array.isArray(P) ||
      P == null ||
      (typeof P == "string" && !P.trim())) &&
      i();
    let L = Number(P);
    Number.isFinite(L) || i();
    let y = qI(T),
      v = qI(A);
    return (
      (!y || !v || y.dimension !== v.dimension) && i("#N/A"),
      y.offset === v.offset
        ? s(L * (y.factor / v.factor))
        : s((L * y.factor + y.offset - v.offset) / v.factor)
    );
  }
  let l = (P) => (Array.isArray(P) ? P.flat(1 / 0) : [P]),
    c = l(r[0]),
    u = l(r[1]);
  c.length !== u.length && i("#N/A");
  let d = c
      .map((P, T) => [P, u[T]])
      .filter((P) => P.every((T) => typeof T == "number")),
    f = d.length;
  (!f && e === "SLOPE" && i("#N/A"),
    f < (e === "STEYX" ? 3 : 2) && i("#DIV/0!"));
  let p = d.map((P) => P[0]),
    h = d.map((P) => P[1]),
    m = sn(p),
    g = sn(h),
    b = h.map((P) => s(P - g)),
    w = p.map((P) => s(P - m)),
    _ = Math.max(...b.map(Math.abs)),
    S = Math.max(...w.map(Math.abs));
  if (((!_ || (!S && (e === "RSQ" || e === "PEARSON"))) && i("#DIV/0!"), !S))
    return 0;
  let O = b.map((P) => P / _),
    E = w.map((P) => P / S),
    R = oi(O.map((P) => P * P)),
    I = oi(E.map((P) => P * P)),
    U = oi(O.map((P, T) => P * E[T]));
  return e === "SLOPE"
    ? s((U / R) * (S / _))
    : e === "PEARSON"
      ? Math.max(-1, Math.min(1, U / Math.sqrt(R * I)))
      : e === "RSQ"
        ? Math.min(1, Math.max(0, (U / R) * (U / I)))
        : s(
            Math.sqrt(
              oi(E.map((P, T) => (P - (U / R) * O[T]) ** 2)) / (f - 2),
            ) * S,
          );
}
var y2 = [
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
  v2 = Symbol("not-statistical-function"),
  Om = (e) => (Array.isArray(e) ? e.flatMap(Om) : [e]),
  oj = (e) => !!e && typeof e == "object" && "ref" in e;
function b2(e) {
  let t = 0,
    n = 0;
  for (let i of e) {
    let r = t + i;
    ((n += Math.abs(t) >= Math.abs(i) ? t - r + i : i - r + t), (t = r));
  }
  return t + n;
}
function HI(e, t, n) {
  if (!y2.includes(e)) return v2;
  let i = (U = "#VALUE!") => {
      throw n.issue(U);
    },
    r = e.startsWith("RANK"),
    o = e.startsWith("PERCENTILE"),
    a = e.startsWith("QUARTILE"),
    s = e === "TRIMMEAN",
    l = e === "PERMUT" || e === "PERMUTATIONA",
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
    ((o || a || s || l) && t.length !== 2) ||
    (r && (t.length < 2 || t.length > 3))
  )
    return i();
  let u = t.map(n.retrieve),
    d = 0,
    f = (U) => {
      if (Array.isArray(U)) {
        for (let P of U) f(P);
        return;
      }
      if (
        (++d > 1e5 && i("#LIMIT!"),
        n.isIssue(U) &&
          ["#LIMIT!", "#CYCLE!", "#BLOCKED!", "#ERROR!"].includes(
            String(U.error),
          ))
      )
        throw U;
    };
  for (let U of u) f(U);
  let p = (U) => (Number.isFinite(U) ? U : i("#NUM!")),
    h = (U, P) => {
      if (U === undefined && P !== undefined) return P;
      if (n.isIssue(U)) throw U;
      if (U == null) return 0;
      if (Array.isArray(U) || (typeof U == "string" && !U.trim())) return i();
      let T = Number(U);
      return Number.isFinite(T) ? T : i();
    };
  if (e === "COUNTA") return u.flatMap(Om).filter((U) => U != null).length;
  if (e === "COUNTBLANK")
    return u.flatMap(Om).filter((U) => U == null || U === "").length;
  let m = r ? 1 : 0;
  if (l) {
    let U = Math.trunc(h(u[0])),
      P = Math.trunc(h(u[1]));
    if (U < 0 || P < 0 || (e === "PERMUT" && P > U) || (!U && P > 0))
      return i("#NUM!");
    if (!P) return 1;
    if (e === "PERMUTATIONA") return p(U ** P);
    if (P > 170) return i("#NUM!");
    let T = 1;
    for (let A = 0; A < P; A++) T = p(T * (U - A));
    return T;
  }
  let g = r || o || a || s ? [u[m]] : u,
    b = [];
  for (let [U, P] of g.entries()) {
    let T = Array.isArray(P) || oj(t[r || o || a || s ? m : U]);
    for (let A of Om(P)) {
      if (n.isIssue(A)) {
        if (e === "COUNT") continue;
        throw A;
      }
      if (typeof A == "number") b.push(p(A));
      else if (c && T && A != null)
        if (typeof A == "boolean") b.push(Number(A));
        else if (typeof A == "string") b.push(0);
        else return i();
      else if (!T && A != null) {
        if (typeof A == "boolean") b.push(Number(A));
        else if (typeof A == "string" && A.trim() && Number.isFinite(Number(A)))
          b.push(Number(A));
        else if (e !== "COUNT") return i();
      }
    }
  }
  if (e === "COUNT") return b.length;
  if (e === "AVERAGE" || e === "AVERAGEA")
    return b.length ? p(sn(b)) : i("#DIV/0!");
  if (e === "MIN" || e === "MINA")
    return b.length ? b.reduce((U, P) => Math.min(U, P)) : 0;
  if (e === "MAX" || e === "MAXA")
    return b.length ? b.reduce((U, P) => Math.max(U, P)) : 0;
  if (e.startsWith("STDEV") || e.startsWith("VAR")) {
    let U = e.endsWith(".S") || e === "STDEVA" || e === "VARA";
    if (b.length < (U ? 2 : 1)) return i("#DIV/0!");
    let P = sn(b),
      T = b.map((v) => v - P),
      A = T.reduce((v, k) => Math.max(v, Math.abs(k)), 0);
    if (
      (Number.isFinite(A)
        ? A && (T = T.map((v) => v / A))
        : ((A = b.reduce((v, k) => Math.max(v, Math.abs(k)), 0)),
          (T = b.map((v) => v / A - P / A))),
      !A)
    )
      return 0;
    let L = b2(T.map((v) => v * v)) / (b.length - (U ? 1 : 0)),
      y = Math.sqrt(L) * A;
    return p(e.startsWith("STDEV") ? y : y * y);
  }
  if (e === "SKEW" || e === "SKEW.P") {
    let U = b.length;
    if (U < 3) return i("#DIV/0!");
    let P = sn(b),
      T = b.reduce((k, x) => Math.max(k, Math.abs(x)), 0);
    if (!T) return i("#DIV/0!");
    let A = b.map((k) => k / T - P / T),
      L = b2(A.map((k) => k * k));
    if (!L) return i("#DIV/0!");
    let v = b2(A.map((k) => k ** 3)) / L ** 1.5;
    return p(
      e === "SKEW.P" ? v * Math.sqrt(U) : (v * U * Math.sqrt(U - 1)) / (U - 2),
    );
  }
  if ((b.sort((U, P) => U - P), s)) {
    let U = h(u[1]);
    if (U < 0 || U >= 1 || !b.length) return i("#NUM!");
    let P = Math.floor((b.length * U) / 2);
    return p(sn(b.slice(P, b.length - P)));
  }
  if (r) {
    let U = h(u[0]),
      P = h(u[2], 0),
      T = b.indexOf(U),
      A = b.lastIndexOf(U);
    return T < 0
      ? i("#N/A")
      : P
        ? 1 + (e === "RANK.AVG" ? (T + A) / 2 : T)
        : b.length - (e === "RANK.AVG" ? (T + A) / 2 : A);
  }
  if (e === "MODE.MULT") {
    let U = new Map(),
      P = 1;
    for (let A of b) {
      let L = (U.get(A) ?? 0) + 1;
      (U.set(A, L), (P = Math.max(P, L)));
    }
    if (P === 1) return i("#N/A");
    let T = [...U].filter(([, A]) => A === P).map(([A]) => [A]);
    return T.length > 1e4 ? i("#LIMIT!") : T;
  }
  if (e === "MODE" || e === "MODE.SNGL") {
    let U = 0,
      P = 1,
      T = 0,
      A;
    for (let L of b)
      ((T = L === A ? T + 1 : 1), T > P && ((P = T), (U = L)), (A = L));
    return P > 1 ? U : i("#N/A");
  }
  if (!b.length) return i("#NUM!");
  if (e === "MEDIAN") {
    let U = Math.floor(b.length / 2);
    return b.length % 2 || b[U - 1] === b[U] ? b[U] : b[U - 1] / 2 + b[U] / 2;
  }
  let w = h(u[1]),
    _ = a ? Math.trunc(w) / 4 : w,
    S = e.endsWith(".EXC");
  if (_ < 0 || _ > 1 || (S && (_ === 0 || _ === 1))) return i("#NUM!");
  let O = S ? _ * (b.length + 1) - 1 : _ * (b.length - 1);
  if (O < 0 || O > b.length - 1) return i("#NUM!");
  let E = Math.floor(O),
    R = Math.ceil(O),
    I = O - E;
  return b[E] === b[R] ? b[E] : p(b[E] * (1 - I) + b[R] * I);
}
var GI = ["LOOKUP", "CHOOSE"],
  w2 = Symbol("not an owned lookup function");
function ZI(e, t, n) {
  if (!["VLOOKUP", "HLOOKUP", "LOOKUP", "CHOOSE"].includes(e)) return w2;
  let i = (m = "#VALUE!") => {
      throw n.issue(m);
    },
    r = (m) => {
      if (n.isIssue(m)) throw m;
      return (Array.isArray(m) && i(), m);
    },
    o = (m) => {
      if (n.isIssue(m)) throw m;
      m === undefined && i();
      let g = Array.isArray(m) ? (m.every(Array.isArray) ? m : [m]) : [[m]];
      return (
        (!g.length ||
          !g[0].length ||
          g.some((b) => b.length !== g[0].length)) &&
          i(),
        g.length * g[0].length > 2e4 && i("#LIMIT!"),
        g
      );
    },
    a = (m) => (m.length !== 1 && m[0].length !== 1 && i(), m.flat());
  if (e === "CHOOSE") {
    (t.length < 2 || t.length > 255) && i();
    let m = r(n.retrieve(t[0]));
    typeof m == "string" && !m.trim() && i();
    let g = Math.trunc(Number(m));
    return ((!Number.isFinite(g) || g < 1 || g >= t.length) && i(), t[g] ?? 0);
  }
  (t.length < (e === "LOOKUP" ? 2 : 3) ||
    t.length > (e === "LOOKUP" ? 3 : 4)) &&
    i();
  let s = r(n.retrieve(t[0]));
  s === undefined && i();
  let l = o(n.retrieve(t[1])),
    c,
    u,
    d = true;
  if (e === "LOOKUP")
    t.length === 3
      ? ((c = a(l)), (u = a(o(n.retrieve(t[2])))), c.length !== u.length && i())
      : l[0].length > l.length
        ? ((c = l[0]), (u = l[l.length - 1]))
        : ((c = l.map((m) => m[0])), (u = l.map((m) => m[m.length - 1])));
  else {
    let m = r(n.retrieve(t[2]));
    typeof m == "string" && !m.trim() && i();
    let g = Math.trunc(Number(m));
    ((!Number.isFinite(g) || g < 1) && i(),
      g > (e === "VLOOKUP" ? l[0].length : l.length) && i("#REF!"));
    let b = t[3] == null ? true : r(n.retrieve(t[3]));
    (typeof b == "string" && !/^(true|false)$/i.test(b) && i(),
      (d = typeof b == "string" ? b.toLowerCase() === "true" : !!b),
      (c = e === "VLOOKUP" ? l.map((w) => w[0]) : l[0]),
      (u = e === "VLOOKUP" ? l.map((w) => w[g - 1]) : l[g - 1]));
  }
  let f = c.find(n.isIssue);
  if (f) throw f;
  let p = (m) => (typeof m == "string" ? m.toLowerCase() : m),
    h = -1;
  if (d) {
    (!["number", "string", "boolean"].includes(typeof s) ||
      c.some((b) => typeof b != typeof s)) &&
      i();
    let m = c.map(p),
      g = p(s);
    (m.some((b, w) => w > 0 && b < m[w - 1]) && i(),
      m.forEach((b, w) => {
        b <= g && (h = w);
      }));
  } else {
    let m = 0,
      g = typeof s == "string" && /[?*~]/.test(s);
    h = c.findIndex((b) =>
      g
        ? typeof b != "string"
          ? false
          : ((m += b.length * s.length),
            m > 1e5 && i("#LIMIT!"),
            n.matches(b, `=${s}`))
        : p(b) === p(s),
    );
  }
  return h < 0 ? i("#N/A") : (u[h] ?? 0);
}
var x2 = Symbol("not an owned logical function");
function KI(e, t, n) {
  if (!["AND", "OR", "XOR", "NOT"].includes(e)) return x2;
  let i = () => {
    throw n.issue("#VALUE!");
  };
  (!t.length || t.length > 255 || (e === "NOT" && t.length !== 1)) && i();
  let r = [],
    o = 0,
    a = (s, l) => {
      if (++o > 1e5) throw n.issue("#LIMIT!");
      if (n.isIssue(s)) throw s;
      if (Array.isArray(s)) (e === "NOT" && i(), s.forEach((c) => a(c, true)));
      else if (typeof s == "boolean" || typeof s == "number") r.push(!!s);
      else if (s == null) l || r.push(false);
      else if (typeof s == "string") {
        if (l) return;
        (/^(true|false)$/i.test(s) || i(), r.push(s.toLowerCase() === "true"));
      } else i();
    };
  return (
    t.forEach((s) =>
      a(
        n.retrieve(s),
        e !== "NOT" && !!s && typeof s == "object" && "ref" in s,
      ),
    ),
    r.length || i(),
    e === "NOT"
      ? !r[0]
      : e === "AND"
        ? r.every(Boolean)
        : e === "OR"
          ? r.some(Boolean)
          : r.filter(Boolean).length % 2 === 1
  );
}
var k2 = [
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
  kd = [
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
  $m = Symbol("not a modern function"),
  _i = (e) => (Array.isArray(e) ? e.flatMap(_i) : [e]),
  Br = (e) => (Array.isArray(e) ? (e.every(Array.isArray) ? e : [e]) : [[e]]),
  rs = (e) =>
    e == null ? "" : typeof e == "boolean" ? (e ? "TRUE" : "FALSE") : String(e),
  Mm = (e, t) =>
    typeof e == "number" && typeof t == "number"
      ? e - t
      : rs(e).toLowerCase().localeCompare(rs(t).toLowerCase()),
  JI = (e, t) =>
    typeof e == "string" && typeof t == "string"
      ? e.toLowerCase() === t.toLowerCase()
      : e === t,
  YI = (e, t, n) => {
    if (e == null || t == null) return e == null ? (t == null ? 0 : 1) : -1;
    let i = (r) => (typeof r == "number" ? 0 : typeof r == "string" ? 1 : 2);
    return (i(e) - i(t) || Mm(e, t)) * n;
  };
function QI(e, t, n) {
  if (!k2.includes(e)) return $m;
  let i = (d = "#VALUE!") => {
      throw n.issue(d);
    },
    r = (d, f) => {
      if (n.isIssue(d)) throw d;
      if (d === undefined && f !== undefined) return f;
      if (Array.isArray(d) || (typeof d == "string" && !d.trim())) return i();
      let p = Number(d);
      return Number.isFinite(p) ? p : i();
    },
    o = (d, f) => Math.trunc(r(d, f)),
    a = (d, f) => {
      ((d < 1 || f < 1) && i("#CALC!"), d * f > 1e4 && i("#LIMIT!"));
    },
    s = (d) => {
      if (n.isIssue(d)) throw d;
      return (
        Array.isArray(d) && i(),
        typeof d == "string" && !/^(true|false)$/i.test(d) && i(),
        typeof d == "string" ? d.toLowerCase() === "true" : !!d
      );
    };
  if (e === "IFS") {
    (t.length < 2 || t.length % 2) && i();
    for (let d = 0; d < t.length; d += 2) {
      let f = n.retrieve(t[d]);
      if ((Array.isArray(f) && i(), s(f))) return n.retrieve(t[d + 1]) ?? 0;
    }
    return i("#N/A");
  }
  if (e === "SWITCH") {
    t.length < 3 && i();
    let d = n.retrieve(t[0]);
    if (n.isIssue(d)) return d;
    Array.isArray(d) && i();
    for (let f = 1; f + 1 < t.length; f += 2) {
      let p = n.retrieve(t[f]);
      if (n.isIssue(p)) return p;
      if ((Array.isArray(p) && i(), JI(d, p))) return n.retrieve(t[f + 1]) ?? 0;
    }
    return t.length % 2 === 0 ? (n.retrieve(t.at(-1)) ?? 0) : i("#N/A");
  }
  let l = t.map(n.retrieve),
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
    (u && (t.length < u[0] || t.length > u[1]) && i(),
    u && e !== "SEQUENCE" && l[0] === undefined && i(),
    u && !["HSTACK", "VSTACK"].includes(e) && n.isIssue(l[0]))
  )
    throw l[0];
  if (!kd.includes(e) && e !== "MAXIFS" && e !== "MINIFS") {
    let d = l.flatMap(_i).find(n.isIssue);
    if (d) return d;
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
    let d = e.startsWith("RANK"),
      f = (
        e === "STDEV.P" || e === "VAR.P" || e === "MODE.SNGL"
          ? l.flatMap(_i)
          : _i(l[d ? 1 : 0])
      )
        .filter((b) => typeof b == "number")
        .sort((b, w) => b - w);
    if ((f.length || i("#NUM!"), d)) {
      let b = r(l[0]),
        w = r(l[2], 0),
        _ = f.indexOf(b),
        S = f.lastIndexOf(b);
      return (
        _ < 0 && i("#N/A"),
        w
          ? 1 + (e === "RANK.AVG" ? (_ + S) / 2 : _)
          : f.length - (e === "RANK.AVG" ? (_ + S) / 2 : S)
      );
    }
    if (e === "MODE.SNGL") {
      let b = new Map();
      f.forEach((_) => b.set(_, (b.get(_) ?? 0) + 1));
      let w = [...b].sort((_, S) => S[1] - _[1] || _[0] - S[0]);
      return w[0][1] > 1 ? w[0][0] : i("#N/A");
    }
    if (e === "STDEV.P" || e === "VAR.P") {
      let b = f.reduce((_, S) => _ + S, 0) / f.length,
        w = f.reduce((_, S) => _ + (S - b) ** 2, 0) / f.length;
      return e === "STDEV.P" ? Math.sqrt(w) : w;
    }
    let p = r(l[1]) / (e.startsWith("QUARTILE") ? 4 : 1),
      h = e.endsWith("EXC"),
      m = h ? p * (f.length + 1) - 1 : p * (f.length - 1);
    (p < 0 || p > 1 || m < 0 || m > f.length - 1) && i("#NUM!");
    let g = Math.floor(m);
    return f[g] + (m - g) * ((f[g + 1] ?? f[g]) - f[g]);
  }
  if (e === "MDETERM" || e === "MINVERSE") {
    t.length !== 1 && i();
    let d = Br(l[0]),
      f = d.length;
    (f > 32 && i("#LIMIT!"), (!f || d.some((g) => g.length !== f)) && i());
    for (let g of d.flat()) {
      if (n.isIssue(g)) return g;
      (typeof g != "number" && i(), Number.isFinite(g) || i("#NUM!"));
    }
    let p = d.map((g, b) => [
        ...g.map((w) => w),
        ...Array.from({ length: f }, (w, _) => (b === _ ? 1 : 0)),
      ]),
      h = 1;
    for (let g = 0; g < f; g++) {
      let b = g;
      for (let _ = g + 1; _ < f; _++)
        Math.abs(p[_][g]) > Math.abs(p[b][g]) && (b = _);
      if (p[b][g] === 0) return e === "MDETERM" ? 0 : i("#NUM!");
      b !== g && (([p[g], p[b]] = [p[b], p[g]]), (h *= -1));
      let w = p[g][g];
      ((h *= w), (p[g] = p[g].map((_) => _ / w)));
      for (let _ = 0; _ < f; _++)
        if (_ !== g) {
          let S = p[_][g];
          p[_] = p[_].map((O, E) => O - S * p[g][E]);
        }
    }
    let m = p.map((g) => g.slice(f));
    return e === "MDETERM"
      ? Number.isFinite(h)
        ? h
        : i("#NUM!")
      : (m.some((g) => g.some((b) => !Number.isFinite(b))) && i("#NUM!"), m);
  }
  if (e === "CONCAT" || e === "TEXTJOIN") {
    e === "TEXTJOIN" && l.length < 3 && i();
    let d = e === "TEXTJOIN" ? rs(l[0]) : "",
      f = e === "TEXTJOIN" && s(l[1]),
      p = l
        .slice(e === "TEXTJOIN" ? 2 : 0)
        .flatMap(_i)
        .filter((h) => !f || (h !== null && h !== ""));
    return (
      p.reduce((h, m) => h + rs(m).length, 0) +
        Math.max(0, p.length - 1) * d.length >
        32767 && i("#LIMIT!"),
      p.map(rs).join(d)
    );
  }
  if (e === "MAXIFS" || e === "MINIFS") {
    (l.length < 3 || l.length % 2 === 0) && i();
    let d = _i(l[0]),
      f = [];
    for (let g = 1; g < l.length; g += 2) {
      let b = _i(l[g]);
      if (
        ((b.length !== d.length ||
          Br(l[g]).length !== Br(l[0]).length ||
          Br(l[g])[0].length !== Br(l[0])[0].length ||
          Array.isArray(l[g + 1])) &&
          i(),
        n.isIssue(l[g + 1]))
      )
        throw l[g + 1];
      f.push([b, l[g + 1]]);
    }
    d.length * f.length > 1e5 && i("#LIMIT!");
    let p = d.filter((g, b) =>
        f.every(([w, _]) => !n.isIssue(w[b]) && n.matches(w[b], _)),
      ),
      h = p.find(n.isIssue);
    if (h) return h;
    let m = p.filter((g) => typeof g == "number");
    return m.length
      ? m.reduce((g, b) => (e === "MAXIFS" ? Math.max(g, b) : Math.min(g, b)))
      : 0;
  }
  if (e === "XMATCH" || e === "XLOOKUP") {
    ((t.length < (e === "XMATCH" ? 2 : 3) ||
      t.length > (e === "XMATCH" ? 4 : 6)) &&
      i(),
      (l[0] === undefined ||
        l[1] === undefined ||
        (e === "XLOOKUP" && l[2] === undefined)) &&
        i());
    let d = l[0],
      f = Br(l[1]);
    if (n.isIssue(d)) return d;
    if (n.isIssue(l[1])) return l[1];
    (Array.isArray(d) ||
      !f.length ||
      !f[0].length ||
      f.some((O) => O.length !== f[0].length) ||
      (f.length > 1 && f[0].length > 1)) &&
      i();
    let p = _i(f),
      h = r(l[e === "XMATCH" ? 2 : 4], 0),
      m = r(l[e === "XMATCH" ? 3 : 5], 1);
    (![0, -1, 1, 2].includes(h) || ![1, -1, 2, -2].includes(m)) && i();
    let g = p.find(n.isIssue);
    if (g) return g;
    let b = [];
    if (e === "XLOOKUP") {
      if (n.isIssue(l[2])) return l[2];
      ((b = Br(l[2])),
        (!b.length ||
          !b[0].length ||
          b.some((O) => O.length !== b[0].length)) &&
          i(),
        (f.length > 1 ? b.length !== p.length : b[0].length !== p.length) &&
          i());
    }
    (Math.abs(m) === 2 || h === -1 || h === 1) &&
      (!["number", "string", "boolean"].includes(typeof d) ||
        p.some((O) => typeof O != typeof d)) &&
      i();
    let w = p.map((O, E) => E);
    (m < 0 && w.reverse(),
      Math.abs(m) === 2 &&
        p.some((O, E) => E > 0 && Mm(O, p[E - 1]) * Math.sign(m) < 0) &&
        i());
    let _ = 0,
      S = w.find((O) =>
        h !== 2 || typeof d != "string"
          ? JI(p[O], d)
          : typeof p[O] != "string"
            ? false
            : ((_ += p[O].length * d.length),
              _ > 1e5 && i("#LIMIT!"),
              n.matches(p[O], `=${d}`)),
      );
    if (S === undefined && (h === -1 || h === 1)) {
      let O = w.filter((E) => Math.sign(Mm(p[E], d)) === h);
      (O.sort((E, R) => Mm(p[E], p[R]) * h), (S = O[0]));
    }
    return S === undefined
      ? e === "XLOOKUP" && l[3] !== undefined
        ? (l[3] ?? 0)
        : i("#N/A")
      : e === "XMATCH"
        ? S + 1
        : f.length > 1
          ? b[S].length === 1
            ? (b[S][0] ?? 0)
            : [b[S]]
          : b.length === 1
            ? (b[0][S] ?? 0)
            : b.map((O) => [O[S]]);
  }
  if (e === "SEQUENCE") {
    l[0] === undefined && l[1] === undefined && i();
    let d = o(l[0], 1),
      f = o(l[1], 1),
      p = r(l[2], 1),
      h = r(l[3], 1);
    return (
      a(d, f),
      Array.from({ length: d }, (m, g) =>
        Array.from({ length: f }, (b, w) => p + (g * f + w) * h),
      )
    );
  }
  if (e === "FILTER") {
    let d = Br(l[0]),
      f = Br(l[1]);
    l[1] === undefined && i();
    let p = _i(f).find(n.isIssue);
    if (p) throw p;
    let h;
    if (f.length === d.length && f[0].length === 1)
      h = d.filter((m, g) => s(f[g][0]));
    else if (f.length === 1 && f[0].length === d[0].length)
      h = d.map((m) => m.filter((g, b) => s(f[0][b])));
    else return i();
    return !h.length || !h[0].length
      ? l[2] !== undefined
        ? Br(l[2])
        : i("#CALC!")
      : h;
  }
  if (e === "TRANSPOSE")
    return Br(l[0])[0].map((d, f) => Br(l[0]).map((p) => p[f]));
  if (e === "SORT" || e === "SORTBY") {
    let d = Br(l[0]).map((b) => [...b]);
    if (e === "SORTBY") {
      let b = Br(l[1]);
      l[1] === undefined && i();
      let w =
          b.length === 1 &&
          b[0].length === d[0].length &&
          (d.length === 1 || b[0].length > 1),
        _ = w ? d[0].length : d.length,
        S = [];
      for (let E = 1; E < l.length; E += 2) {
        l[E] === undefined && i();
        let R = Br(l[E]),
          I = r(l[E + 1], 1);
        ((w
          ? R.length !== 1 || R[0].length !== _
          : R.length !== _ || R.some((T) => T.length !== 1)) ||
          ![-1, 1].includes(I)) &&
          i();
        let U = _i(R),
          P = U.find(n.isIssue);
        if (P) throw P;
        S.push({ data: U, order: I });
      }
      let O = Array.from({ length: _ }, (E, R) => R).sort((E, R) => {
        for (let I of S) {
          let U = YI(I.data[E], I.data[R], I.order);
          if (U) return U;
        }
        return E - R;
      });
      return w ? d.map((E) => O.map((R) => E[R])) : O.map((E) => d[E]);
    }
    let f = o(l[1], 1) - 1,
      p = r(l[2], 1),
      h = l[3] === undefined ? false : s(l[3]);
    [-1, 1].includes(p) || i();
    let m = h ? d[0].map((b, w) => d.map((_) => _[w])) : d;
    (f < 0 || f >= m[0].length) && i();
    let g = m.map((b) => b[f]).find(n.isIssue);
    if (g) throw g;
    return (
      m.sort((b, w) => YI(b[f], w[f], p)),
      h ? m[0].map((b, w) => m.map((_) => _[w])) : m
    );
  }
  if (e === "UNIQUE") {
    let d = Br(l[0]),
      f = l[1] === undefined ? false : s(l[1]),
      p = l[2] === undefined ? false : s(l[2]),
      h = f ? d[0].map((_, S) => d.map((O) => O[S])) : d,
      m = (_) =>
        JSON.stringify(
          _.map((S) => (typeof S == "string" ? S.toLowerCase() : (S ?? 0))),
        ),
      g = new Map();
    h.forEach((_) => g.set(m(_), (g.get(m(_)) ?? 0) + 1));
    let b = new Set(),
      w = h.filter((_) => {
        let S = m(_);
        return b.has(S) || (p && g.get(S) !== 1) ? false : (b.add(S), true);
      });
    return (
      w.length || i("#CALC!"),
      f ? w[0].map((_, S) => w.map((O) => O[S])) : w
    );
  }
  if (e === "TAKE" || e === "DROP") {
    let d = Br(l[0]),
      f = l[1] === undefined ? undefined : o(l[1]),
      p = l[2] === undefined ? undefined : o(l[2]),
      h = (g, b) =>
        b === undefined
          ? g
          : e === "TAKE"
            ? b > 0
              ? g.slice(0, b)
              : g.slice(b)
            : b >= 0
              ? g.slice(b)
              : g.slice(0, b);
    (f === undefined && p === undefined && i(),
      (f === 0 || p === 0) && i("#CALC!"));
    let m = h(d, f).map((g) => h(g, p));
    return ((!m.length || !m[0].length) && i("#CALC!"), m);
  }
  if (e === "CHOOSECOLS" || e === "CHOOSEROWS") {
    let d = Br(l[0]),
      f = e === "CHOOSECOLS" ? d[0].length : d.length,
      p = l
        .slice(1)
        .flatMap(_i)
        .map((h) => {
          let m = o(h),
            g = m > 0 ? m - 1 : f + m;
          return ((!m || g < 0 || g >= f) && i(), g);
        });
    return (
      p.length || i(),
      e === "CHOOSECOLS" ? d.map((h) => p.map((m) => h[m])) : p.map((h) => d[h])
    );
  }
  if (e === "HSTACK" || e === "VSTACK") {
    let d = l.map(Br),
      f =
        e === "VSTACK"
          ? d.reduce((h, m) => h + m.length, 0)
          : Math.max(...d.map((h) => h.length)),
      p =
        e === "HSTACK"
          ? d.reduce((h, m) => h + m[0].length, 0)
          : Math.max(...d.map((h) => h[0].length));
    return (
      a(f, p),
      e === "VSTACK"
        ? d.flatMap((h) =>
            h.map((m) =>
              Array.from({ length: p }, (g, b) =>
                b < m.length ? m[b] : n.issue("#N/A"),
              ),
            ),
          )
        : Array.from({ length: f }, (h, m) =>
            d.flatMap((g) =>
              Array.from({ length: g[0].length }, (b, w) =>
                m < g.length ? g[m][w] : n.issue("#N/A"),
              ),
            ),
          )
    );
  }
  if (e === "TOCOL" || e === "TOROW") {
    let d = Br(l[0]),
      f = o(l[1], 0),
      p = l[2] === undefined ? false : s(l[2]);
    [0, 1, 2, 3].includes(f) || i();
    let h = (p ? d[0].flatMap((m, g) => d.map((b) => b[g])) : d.flat()).filter(
      (m) =>
        !((f === 1 || f === 3) && m === null) &&
        !((f === 2 || f === 3) && n.isIssue(m)),
    );
    return (
      h.length || i("#CALC!"),
      e === "TOCOL" ? h.map((m) => [m ?? 0]) : [h.map((m) => m ?? 0)]
    );
  }
  if (e === "EXPAND") {
    let d = Br(l[0]),
      f = o(l[1], d.length),
      p = o(l[2], d[0].length);
    return (
      a(f, p),
      (f < d.length || p < d[0].length) && i(),
      Array.from({ length: f }, (h, m) =>
        Array.from({ length: p }, (g, b) =>
          m < d.length && b < d[m].length ? d[m][b] : (l[3] ?? n.issue("#N/A")),
        ),
      )
    );
  }
  if (e === "TEXTBEFORE" || e === "TEXTAFTER") {
    let d = rs(l[0]),
      f = rs(l[1]),
      p = o(l[2], 1),
      h = o(l[3], 0) === 1,
      m = o(l[4], 0) === 1;
    (!p || !f) && i();
    let g = h ? d.toLowerCase() : d,
      b = h ? f.toLowerCase() : f,
      w = [];
    for (let S = 0; S <= g.length - b.length;) {
      let O = g.indexOf(b, S);
      if (O < 0) break;
      (w.push(O), (S = O + b.length));
    }
    let _ = w[p > 0 ? p - 1 : w.length + p];
    return _ === undefined
      ? m && Math.abs(p) === w.length + 1
        ? e === "TEXTBEFORE"
          ? p > 0
            ? d
            : ""
          : p > 0
            ? ""
            : d
        : l[5] !== undefined
          ? l[5]
          : i("#N/A")
      : e === "TEXTBEFORE"
        ? d.slice(0, _)
        : d.slice(_ + f.length);
  }
  if (e === "TEXTSPLIT") {
    (l.length < 2 || l.length > 6) && i();
    let d = l.slice(0, 5).flatMap(_i).find(n.isIssue);
    if (d) return d;
    let f = rs(l[0]),
      p = l[3] === undefined ? false : s(l[3]),
      h = r(l[4], 0);
    (Array.isArray(l[0]) || ![0, 1].includes(h)) && i();
    let m = (O) =>
        _i(O)
          .filter((E) => E != null)
          .map((E) => {
            if (n.isIssue(E)) throw E;
            return rs(E);
          })
          .filter(Boolean),
      g = m(l[1]),
      b = m(l[2]);
    (!g.length && !b.length && i(),
      (g.length + b.length > 100 || [...g, ...b].join("").length > 1e4) &&
        i("#LIMIT!"));
    let w = (O, E) =>
        E.length
          ? O.split(
              new RegExp(
                E.map((R) => R.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")).join(
                  "|",
                ),
                h ? "i" : "",
              ),
            ).filter((R) => !p || R !== "")
          : [O],
      _ = w(f, b).map((O) => w(O, g)),
      S = Math.max(..._.map((O) => O.length));
    return (
      a(_.length, S),
      _.map((O) =>
        Array.from({ length: S }, (E, R) => O[R] ?? l[5] ?? n.issue("#N/A")),
      )
    );
  }
  return $m;
}
function _d(e, t) {
  if (e === undefined) return [];
  if (!Array.isArray(e) || e.length > 1e3)
    throw new Error("A workbook may contain at most 1,000 defined names.");
  let n = new Set();
  return e.map((i) => {
    if (!i || typeof i != "object" || Array.isArray(i))
      throw new Error("Invalid defined name.");
    let r = i;
    if (
      Object.keys(r).some(
        (a) => !["name", "refersTo", "scopeSheetId", "comment"].includes(a),
      )
    )
      throw new Error("Unknown defined-name field.");
    if (
      typeof r.name != "string" ||
      !/^[A-Za-z_\\][A-Za-z0-9_.\\]{0,254}$/.test(r.name) ||
      /^(?:[A-Za-z]{1,3}[1-9]\d*|R\d+C\d+|R|C)$/i.test(r.name)
    )
      throw new Error(
        "Names must begin with a letter or underscore and cannot be cell addresses.",
      );
    if (
      typeof r.refersTo != "string" ||
      !r.refersTo.startsWith("=") ||
      r.refersTo.length > 4096 ||
      r.refersTo.length < 2
    )
      throw new Error(
        "A name must refer to a formula beginning with = (up to 4,096 characters).",
      );
    if (
      r.scopeSheetId !== undefined &&
      (typeof r.scopeSheetId != "string" ||
        !t.sheets.some((a) => a.id === r.scopeSheetId))
    )
      throw new Error("Defined-name worksheet scope was not found.");
    if (
      r.comment !== undefined &&
      (typeof r.comment != "string" || r.comment.length > 1e3)
    )
      throw new Error(
        "Defined-name comments may contain at most 1,000 characters.",
      );
    let o = `${r.scopeSheetId ?? ""}:${r.name.toLowerCase()}`;
    if (n.has(o)) throw new Error("A name already exists in this scope.");
    return (
      n.add(o),
      {
        name: r.name,
        refersTo: r.refersTo,
        ...(r.scopeSheetId ? { scopeSheetId: r.scopeSheetId } : {}),
        ...(r.comment ? { comment: r.comment } : {}),
      }
    );
  });
}
var sj = (e) => `${e.scopeSheetId ?? ""}:${e.name.toLowerCase()}`,
  aj = (e) => `'${e.replace(/'/g, "''")}'!`;
function lj(e, t = false) {
  let n = [],
    i = 0,
    r = (a) => {
      if (t) throw Error(a);
    },
    o = () => {
      let a = 0;
      do (e[i] === "[" && a++, e[i] === "]" && a--, i++);
      while (i < e.length && a);
      a && r("Unbalanced table references cannot be safely migrated.");
    };
  for (; i < e.length;) {
    let a = i,
      s = e[i];
    if (s === '"') {
      i++;
      let h = false;
      for (; i < e.length;)
        if (e[i++] === '"')
          if (e[i] === '"') i++;
          else {
            h = true;
            break;
          }
      h || r("An unterminated formula string cannot be safely migrated.");
      continue;
    }
    if (s === "#") {
      let h =
        /^#(?:REF!|NAME\?|VALUE!|DIV\/0!|N\/A|NUM!|NULL!|SPILL!|CALC!|CYCLE!|LIMIT!|BLOCKED!|ERROR!)/i.exec(
          e.slice(i),
        );
      if (h) {
        i += h[0].length;
        continue;
      }
      i++;
      continue;
    }
    let l =
      /^(?:'(?:[^']|'')+'|[A-Za-z0-9_.]+):(?:'(?:[^']|'')+'|[A-Za-z0-9_.]+)!/.exec(
        e.slice(i),
      );
    if (l) {
      (r("Three-dimensional references cannot be safely migrated."),
        (i += l[0].length));
      continue;
    }
    let c;
    if (s === "'") {
      let h = /^'((?:[^']|'')+)'!/.exec(e.slice(i));
      if (!h) {
        for (
          r("An unsupported quoted reference cannot be safely migrated."), i++;
          i < e.length && e[i++] !== "'";
        );
        continue;
      }
      ((c = h[1].replace(/''/g, "'")),
        (i += h[0].length),
        /[\[\]:|]/.test(c) &&
          r(
            "External or three-dimensional references cannot be safely migrated.",
          ));
    } else {
      let h = /^([A-Za-z0-9_\\][A-Za-z0-9_.\\]*)!/.exec(e.slice(i));
      h && ((c = h[1]), (i += h[0].length));
    }
    if (c === undefined && s === "[") {
      (e.startsWith("[@", i) ||
        r(
          "External or unqualified bracket references cannot be safely migrated.",
        ),
        o());
      continue;
    }
    let u =
      /^(?:\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}|\$?[1-9]\d*:\$?[1-9]\d*)(?![A-Za-z0-9_.\\])/.exec(
        e.slice(i),
      );
    if (u) {
      ((i += u[0].length),
        n.push({
          start: a,
          end: i,
          text: e.slice(a, i),
          name: u[0],
          ...(c ? { qualifier: c } : {}),
          cell: true,
        }));
      continue;
    }
    if (c === undefined && /[0-9.]/.test(s)) {
      let h = /^(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?/.exec(e.slice(i));
      if (h) {
        i += h[0].length;
        continue;
      }
    }
    let d =
      /^(?:\$?[A-Za-z]{1,3}\$?[1-9]\d*(?![A-Za-z0-9_.\\])|[A-Za-z_\\][A-Za-z0-9_.\\]*)/.exec(
        e.slice(i),
      );
    if (!d) {
      (c && r("The qualified reference cannot be safely migrated."),
        s === "!" && r("The reference qualifier cannot be safely migrated."),
        s === "@" && e[i + 1] === "[" ? (i++, o()) : i++);
      continue;
    }
    let f = d[0];
    if (((i += f.length), e[i] === "[")) {
      o();
      continue;
    }
    let p = /^\s*\(/.test(e.slice(i));
    (p &&
      /^(?:_xlfn\.)?(?:INDIRECT|EVALUATE|GET\.CELL)$/i.test(f) &&
      r(
        `${f} uses dynamic or locally bound names and cannot be safely migrated.`,
      ),
      !(a > 0 && /[A-Za-z0-9_.$\\]/.test(e[a - 1])) &&
        n.push({
          start: a,
          end: i,
          text: e.slice(a, i),
          name: f,
          ...(p ? { call: true } : {}),
          ...(c ? { qualifier: c } : {}),
          ...(/^\$?[A-Za-z]{1,3}\$?[1-9]\d*$/.test(f) ? { cell: true } : {}),
        }));
  }
  return n;
}
function e5(e, t = false) {
  let n = lj(e, t);
  if (!/\b(?:_xlfn\.)?(?:LET|LAMBDA)\s*\(/i.test(e)) return n;
  let i = new Map();
  try {
    Fo(Vl(e), (r, o, a, s) => {
      i.set(r.start, {
        bound: new Set(o),
        excluded: s || (!r.qualifier && o.has(mo(r.name))),
      });
    });
  } catch (r) {
    if (t) throw r;
    return n;
  }
  return n
    .filter((r) => !i.get(r.start)?.excluded)
    .map((r) => ({ ...r, bound: i.get(r.start)?.bound }));
}
function cj(e, t, n) {
  if (e.cell || (e.call && !e.qualifier && Pm.includes(Du(e.name)))) return;
  let i = e.qualifier
    ? t.sheets.find((r) => r.name.toLowerCase() === e.qualifier.toLowerCase())
        ?.id
    : n;
  if (i)
    return (
      t.definedNames?.find(
        (r) =>
          r.scopeSheetId === i && r.name.toLowerCase() === e.name.toLowerCase(),
      ) ??
      t.definedNames?.find(
        (r) => !r.scopeSheetId && r.name.toLowerCase() === e.name.toLowerCase(),
      )
    );
}
function t5(e, t, n) {
  let i = "",
    r = 0;
  for (let o of t) ((i += e.slice(r, o.start) + n(o)), (r = o.end));
  return i + e.slice(r);
}
function Ed(e, t, n, i = [], r) {
  if (i.length > 32) throw new Error("#LIMIT!");
  let o = t5(e, e5(e), (a) => {
    if (a.call) return a.text;
    let s = cj(a, t, n);
    if (!s || v6(s.refersTo)) return a.text;
    let l = sj(s);
    if (i.includes(l)) throw new Error("#CYCLE!");
    let c =
      s.scopeSheetId ??
      (a.qualifier
        ? t.sheets.find(
            (d) => d.name.toLowerCase() === a.qualifier.toLowerCase(),
          ).id
        : n);
    if (r && !r(c)) throw new Error("#BLOCKED!");
    let u = Ed(s.refersTo.slice(1), t, c, [...i, l], r);
    if (c !== n) {
      let d = t.sheets.find((p) => p.id === c),
        f = e5(u);
      u = t5(u, f, (p) => {
        let h = f.indexOf(p),
          m = f[h - 1],
          g = m?.cell && u.slice(m.end, p.start).trim() === ":";
        return p.cell && !p.qualifier && !g ? aj(d.name) + p.text : p.text;
      });
    }
    return `(${u})`;
  });
  if (o.length > 16384) throw new Error("#LIMIT!");
  return o;
}
function Co(e) {
  let t = "",
    n = 0;
  for (; n < e.length;) {
    let i = e[n];
    if (i === '"' || i === "'") {
      let o = n++,
        a = i;
      for (; n < e.length;)
        if (e[n++] === a)
          if (e[n] === a) n++;
          else break;
      t += e.slice(o, n);
      continue;
    }
    let r = /^(?:[A-Za-z_][A-Za-z0-9_.]*\[|\[@)/.exec(e.slice(n));
    if (r && (n === 0 || !/[A-Za-z0-9_.$]/.test(e[n - 1]))) {
      let o = n;
      n += r[0].length;
      let a = 1;
      for (; n < e.length && a;)
        (e[n] === "[" && a++, e[n] === "]" && a--, n++);
      let s = e.slice(o, n);
      !a && !/[!|]/.test(s) && !/^[A-Za-z0-9_!]/.test(e.slice(n))
        ? (t += "0")
        : (t += s);
      continue;
    }
    t += e[n++];
  }
  return t;
}
var uj = new Set([
    "_xlnm.print_area",
    "_xlnm.print_titles",
    "_xlnm._filterdatabase",
  ]),
  Sd = "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
  Dt = (e) => {
    throw Error(
      `XLSX defined names: ${e} Import/export stopped to prevent changed formula bindings.`,
    );
  },
  _2 =
    /[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]|[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/;
function n5(e) {
  return e.includes("<") || e.includes("]]>") || _2.test(e)
    ? Dt("Malformed XML text.")
    : e.replace(/&[^;\s<]*;|&/g, (t) => {
        let n = {
          "&amp;": "&",
          "&lt;": "<",
          "&gt;": ">",
          "&quot;": '"',
          "&apos;": "'",
        };
        if (n[t]) return n[t];
        if (!/^&#(?:[0-9]+|x[0-9a-fA-F]+);$/.test(t))
          return Dt("Unknown or malformed XML entity.");
        let i = t.startsWith("&#x"),
          r = Number.parseInt(t.slice(i ? 3 : 2, -1), i ? 16 : 10);
        if (
          !Number.isSafeInteger(r) ||
          r > 1114111 ||
          r < 0 ||
          (r >= 55296 && r <= 57343)
        )
          return Dt("Invalid XML character entity.");
        let o = String.fromCodePoint(r);
        return _2.test(o) ? Dt("Invalid XML character entity.") : o;
      });
}
function r5(e, t = false) {
  return _2.test(e)
    ? Dt("A name contains an invalid XML character.")
    : e.replace(
        /[&<>"'\r\n\t]/g,
        (n) =>
          ({
            "&": "&amp;",
            "<": "&lt;",
            ">": "&gt;",
            '"': "&quot;",
            "'": "&apos;",
          })[n] ?? (t ? `&#${n.charCodeAt(0)};` : n === "\r" ? "&#13;" : n),
      );
}
function E2(e) {
  let t = Object.create(null),
    n = e;
  for (; n.trim();) {
    let i = /^\s+([A-Za-z_][A-Za-z0-9_.:-]*)\s*=\s*(["'])([\s\S]*?)\2/.exec(n);
    if (!i || Object.hasOwn(t, i[1]))
      return Dt("Malformed or duplicate defined-name XML attribute.");
    ((t[i[1]] = n5(i[3].replace(/\r\n?|\n|\t/g, " "))),
      (n = n.slice(i[0].length)));
  }
  return t;
}
function i5(e) {
  if (/<!DOCTYPE|<!ENTITY/i.test(e))
    return Dt("DTD and entity declarations are not supported.");
  let t = [
    ...e.matchAll(
      /<((?:[A-Za-z_][\w.-]*:)?definedNames)\b((?:[^"'<>]|"[^"]*"|'[^']*')*?)(?:\/>|>([\s\S]*?)<\/\1\s*>)/g,
    ),
  ];
  if (t.length > 1) return Dt("Duplicate definedNames sections.");
  let n = t[0],
    i = n ? e.slice(0, n.index) + e.slice(n.index + n[0].length) : e;
  if (/<\/?(?:[A-Za-z_][\w.-]*:)?definedNames?\b/.test(i))
    return Dt("Malformed or misplaced defined-name elements.");
  if (!n) return;
  let r = E2(n[2]),
    o =
      /<(?:[A-Za-z_][\w.-]*:)?workbook\b((?:[^"'<>]|"[^"]*"|'[^']*')*?)>/.exec(
        e,
      ),
    a = { ...(o ? E2(o[1]) : {}), ...r },
    s = n[1].slice(0, -12);
  if (a[s ? `xmlns:${s.slice(0, -1)}` : "xmlns"] !== Sd)
    return Dt("Unsupported or undeclared definedNames namespace.");
  for (let l of Object.keys(r))
    if (!l.startsWith("xmlns") || r[l] !== Sd)
      return Dt("Unsupported definedNames namespace or attribute.");
  return {
    start: n.index,
    end: n.index + n[0].length,
    prefix: n[1].slice(0, -12),
    body: n[3] ?? "",
    namespaces: a,
  };
}
function Cd(e) {
  if (!e.startsWith("=") || e.length < 2 || e.length > 4096)
    return Dt(
      "A definition must contain a formula of 1\u20134,095 characters after =.",
    );
  let t = 0,
    n = 0,
    i = 0,
    r = 1,
    o = "";
  for (; r < e.length;) {
    let s = e[r];
    if (!n && (s === '"' || s === "'")) {
      let l = s,
        c = r++,
        u = false;
      for (; r < e.length;)
        if (e[r++] === l)
          if (e[r] === l) r++;
          else {
            u = true;
            break;
          }
      if (!u) return Dt("Unterminated formula string or worksheet qualifier.");
      if (l === "'") {
        if (e[r] !== "!") return Dt("Malformed quoted worksheet qualifier.");
        o += e.slice(c, r);
      } else o += '""';
      continue;
    }
    if ((s === "[" && n++, s === "]" && --n < 0))
      return Dt("Unbalanced bracket reference.");
    if (!n) {
      if ((s === "{" && i++, s === "}" && --i < 0))
        return Dt("Unbalanced array braces.");
      if ((s === "(" && t++, s === ")" && --t < 0))
        return Dt("Unbalanced formula parentheses.");
    }
    ((o += s), r++);
  }
  if (t || n || i)
    return Dt(
      "Unbalanced formula parentheses, array braces or bracket reference.",
    );
  let a = Co(o);
  if (
    /[\[\]|]/.test(a) ||
    /(?:https?:|file:|javascript:|data:|\\\\)/i.test(a) ||
    /\b(?:_xlfn\.)?(?:WEBSERVICE|HYPERLINK|IMAGE|RTD|DDE|CALL|EXEC|EXECUTE|REGISTER(?:\.ID)?|RUN|SQL\.REQUEST|STOCKHISTORY|EVALUATE|GET\.CELL)\s*\(/i.test(
      a,
    )
  )
    return Dt("External or executable name formulas are not supported.");
  if (
    /(?:'(?:[^']|'')+'|[A-Za-z0-9_.]+):(?:'(?:[^']|'')+'|[A-Za-z0-9_.]+)!/.test(
      a,
    )
  )
    return Dt("Three-dimensional name references are not supported.");
  if (!o.trim()) return Dt("An empty formula is not supported.");
}
function dj(e, t) {
  let n = Co(e.replace(/"(?:[^"]|"")*"/g, '""'));
  for (let i of n.matchAll(
    /(?:'((?:[^']|'')+)'|([A-Za-z_\\][A-Za-z0-9_.\\]*))!/g,
  )) {
    let r = i[1]?.replace(/''/g, "'") ?? i[2];
    if (!t.sheets.some((o) => o.name.toLowerCase() === r.toLowerCase()))
      return Dt(`A name refers to an unknown worksheet "${r}".`);
  }
}
function fj(e) {
  let t = 0;
  for (let { refersTo } of e) {
    let i = Co(
      refersTo.replace(/"(?:[^"]|"")*"/g, '""').replace(/'(?:[^']|'')+'!/g, ""),
    );
    for (let r of i.matchAll(
      /(?:^|[^A-Za-z0-9_.])((?:\$?[A-Za-z]{1,3}\$?[1-9]\d*)(?::\$?[A-Za-z]{1,3}\$?[1-9]\d*)?)(?![A-Za-z0-9_.])/g,
    )) {
      let o = Qe(r[1]);
      if (
        ((t += (o.lastRow - o.firstRow + 1) * (o.lastCol - o.firstCol + 1)),
        t > Ne.maxCells)
      )
        throw Error(
          "Named ranges exceed the 100,000-cell XLSX import/export processing budget. Use native JSON for larger definitions.",
        );
    }
  }
}
function o5(e) {
  let t = i5(e),
    n = { names: [], warnings: [], builtinXml: "" };
  if (!t) return n;
  let i = t.body,
    r = 0,
    o = new Set(),
    a = [...e.matchAll(/<(?:[A-Za-z_][\w.-]*:)?sheet\b[^>]*>/g)].length;
  for (; i.trim();) {
    let s =
      /^\s*<((?:[A-Za-z_][\w.-]*:)?definedName)\b((?:[^"'<>]|"[^"]*"|'[^']*')*?)(?:\/>|>([^<]*)<\/\1\s*>)/.exec(
        i,
      );
    if (!s)
      return Dt(
        "Malformed definedName content; only text definitions are supported.",
      );
    let l = E2(s[2]),
      c = l.name ?? "",
      u = l.localSheetId,
      d = s[1].slice(0, -11);
    if (
      { ...t.namespaces, ...l }[d ? `xmlns:${d.slice(0, -1)}` : "xmlns"] !== Sd
    )
      return Dt("Unsupported or undeclared definedName namespace.");
    for (let m of Object.keys(l))
      if (m.startsWith("xmlns") && l[m] !== Sd)
        return Dt("Unsupported definedName namespace.");
    if (
      u !== undefined &&
      (!/^\d+$/.test(u) || !Number.isSafeInteger(Number(u)))
    )
      return Dt(`Name "${c}" has an invalid localSheetId.`);
    let p = `=${n5(
      (s[3] ?? "").replace(
        /\r\n?/g,
        `
`,
      ),
    ).replace(/^=/, "")}`;
    if ((Cd(p), u !== undefined && Number(u) >= a))
      return Dt(`Name "${c}" has an out-of-range localSheetId.`);
    let h = `${u === undefined ? "workbook" : Number(u)}:${c.toLowerCase()}`;
    if (o.has(h)) return Dt(`Duplicate name "${c}" in the same scope.`);
    if ((o.add(h), uj.has(c.toLowerCase())))
      n.builtinXml += `<definedName${Object.entries(l)
        .filter(([m]) => !m.startsWith("xmlns"))
        .map(([m, g]) => ` ${m}="${r5(g, true)}"`)
        .join("")}>${r5(p.slice(1))}</definedName>`;
    else if (c.toLowerCase().startsWith("_xlnm.")) {
      if (++r > 1e3)
        return Dt("The workbook exceeds the 1,000-defined-name limit.");
      n.warnings.push(
        `Built-in defined name "${c}" is not supported and was omitted. Review its source feature before relying on imported formulas.`,
      );
    } else {
      if (++r > 1e3)
        return Dt("The workbook exceeds the 1,000-defined-name limit.");
      for (let g of ["function", "vbProcedure", "xlm"])
        if (l[g] !== undefined && !["0", "false"].includes(l[g]))
          return Dt(`Name "${c}" has an executable ${g} attribute.`);
      let m = Object.keys(l).filter(
        (g) =>
          !["name", "localSheetId", "comment"].includes(g) &&
          !g.startsWith("xmlns"),
      );
      (m.length &&
        n.warnings.push(
          `Defined name "${c}" retained its definition and scope; unsupported attributes (${m.join(", ")}) were omitted. Hidden names are visible in Name Manager.`,
        ),
        n.names.push({
          name: c,
          refersTo: p,
          ...(u !== undefined ? { localSheetId: Number(u) } : {}),
          ...(l.comment !== undefined ? { comment: l.comment } : {}),
        }));
    }
    i = i.slice(s[0].length);
  }
  fj(n.names);
  try {
    _d(
      n.names.map((s) => ({
        name: s.name,
        refersTo: s.refersTo,
        ...(s.localSheetId !== undefined
          ? { scopeSheetId: String(s.localSheetId) }
          : {}),
        ...(s.comment !== undefined ? { comment: s.comment } : {}),
      })),
      {
        sheets: [
          ...new Set(
            n.names.flatMap((s) =>
              s.localSheetId === undefined ? [] : [s.localSheetId],
            ),
          ),
        ].map((s) => ({ id: String(s) })),
      },
    );
  } catch (s) {
    return Dt(s.message);
  }
  return n;
}
function s5(e, t, n) {
  if (new Set(n.map((r) => r.toLowerCase())).size !== n.length)
    return Dt("Duplicate worksheet names make name scopes ambiguous.");
  let i = e.map((r) => {
    let o = r.localSheetId === undefined ? undefined : n[r.localSheetId],
      a = o === undefined ? undefined : t.sheets.find((s) => s.name === o)?.id;
    return r.localSheetId !== undefined && !a
      ? Dt(`Name "${r.name}" has an invalid or unsupported worksheet scope.`)
      : {
          name: r.name,
          refersTo: r.refersTo,
          ...(a ? { scopeSheetId: a } : {}),
          ...(r.comment !== undefined ? { comment: r.comment } : {}),
        };
  });
  for (let r of i) dj(r.refersTo, t);
  try {
    return _d(i, t);
  } catch (r) {
    return Dt(r.message);
  }
}
function a5(e, t) {
  let n = i5(e),
    i = t ? `<definedNames xmlns="${Sd}">${t}</definedNames>` : "";
  if (n) return e.slice(0, n.start) + i + e.slice(n.end);
  if (!t) return e;
  let r =
    /<(?:[A-Za-z_][\w.-]*:)?(?:calcPr|oleSize|customWorkbookViews|pivotCaches|smartTagPr|smartTagTypes|webPublishing|fileRecoveryPr|webPublishObjects|extLst)\b|<\/(?:[A-Za-z_][\w.-]*:)?workbook\s*>/.exec(
      e,
    );
  return r
    ? e.slice(0, r.index) + i + e.slice(r.index)
    : Dt("The workbook XML has no supported insertion point.");
}
var l5 = ["LET", "LAMBDA"],
  hc = { depth: 64, calls: 1024, work: 1e6, retainedValues: 1e5 },
  S2 = (e) =>
    !!e && typeof e == "object" && "kind" in e && e.kind === "formula-closure",
  Dm = "__LGVALUE";
function c5(e, t, n, i) {
  let r = 0,
    o = 0,
    a = 0,
    s = new Map(),
    l = new Set(),
    c = (S, O) => {
      throw (i.diagnostic(S, O), i.issue(S));
    },
    u = (S, O) => {
      ((r += S),
        (O > hc.depth || r > hc.work) &&
          c(
            "#LIMIT!",
            `Reusable formula evaluation exceeds ${hc.depth} levels or ${hc.work.toLocaleString()} work units.`,
          ));
    },
    d = (S) => {
      let O = Array.isArray(S)
        ? S.reduce((E, R) => E + (Array.isArray(R) ? R.length : 1), 0)
        : 1;
      return (
        (a += O),
        a > hc.retainedValues &&
          c(
            "#LIMIT!",
            "Reusable formula intermediate values exceed 100,000 cells.",
          ),
        S
      );
    },
    f = (S) => {
      let O = s.get(S);
      if (!O) {
        u(S.length, 0);
        try {
          try {
            Cd(S.startsWith("=") ? S : `=${S}`);
          } catch (E) {
            let R = E instanceof Error ? E.message : "Unsafe formula.";
            c(
              /External|executable|Three-dimensional/i.test(R)
                ? "#BLOCKED!"
                : "#VALUE!",
              R,
            );
          }
          ((O = Vl(S)),
            Fo(O, (E, R, I, U) => {
              (!U &&
                /^__LG/i.test(E.name) &&
                c(
                  "#NAME?",
                  "Internal calculation names cannot be called from a workbook.",
                ),
                /^(?:_xlfn\.)?(?:WEBSERVICE|HYPERLINK|IMAGE|RTD|DDE|CALL|EXEC|EXECUTE|REGISTER(?:\.ID)?|RUN|SQL\.REQUEST|STOCKHISTORY|EVALUATE|GET\.CELL)$/i.test(
                  E.name,
                ) &&
                  c(
                    "#BLOCKED!",
                    "External or executable custom functions are not supported.",
                  ));
            }));
        } catch (E) {
          if (i.isIssue(E)) throw E;
          c(
            E instanceof Error && E.message === "#LIMIT!"
              ? "#LIMIT!"
              : "#VALUE!",
            E instanceof Error ? E.message : "Invalid reusable formula.",
          );
        }
        s.set(S, O);
      }
      return O;
    },
    p = (S, O) => {
      let E = S.qualifier
        ? t.sheets.find(
            (U) => U.name.toLowerCase() === S.qualifier.toLowerCase(),
          )?.id
        : O;
      if (!E) return;
      let R = mo(S.name),
        I =
          t.definedNames?.find(
            (U) => U.scopeSheetId === E && U.name.toLowerCase() === R,
          ) ??
          t.definedNames?.find(
            (U) => !U.scopeSheetId && U.name.toLowerCase() === R,
          );
      return I ? { found: I, sheetId: I.scopeSheetId ?? E } : undefined;
    },
    h = (S, O, E) => {
      try {
        return i.parse(S, O, E);
      } catch (R) {
        let I =
          R && typeof R == "object" && "details" in R ? R.details : undefined;
        return i.isIssue(I) ? I : i.isIssue(R) ? R : i.issue("#ERROR!");
      }
    },
    m = (S, O, E, R, I) => {
      let U = mo(S.name);
      if (!S.qualifier && E.has(U)) return E.get(U);
      if (!S.qualifier && /^(?:TRUE|FALSE)$/i.test(S.name)) return U === "true";
      let P = p(S, R);
      if (!P)
        return c(
          "#NAME?",
          `Unknown name or unsupported custom function: ${O.slice(S.start, S.end)}.`,
        );
      let T = `${P.sheetId}:${P.found.name.toLowerCase()}`;
      if (i.allowName && !i.allowName(P.sheetId))
        return c(
          "#BLOCKED!",
          "The defined name is outside the authorized inspection context.",
        );
      if (l.has(T))
        return c("#CYCLE!", `Circular defined name: ${P.found.name}.`);
      l.add(T);
      try {
        return _(
          f(P.found.refersTo),
          P.found.refersTo,
          new Map(),
          P.sheetId,
          I + 1,
        );
      } finally {
        l.delete(T);
      }
    },
    g = (S, O, E) => {
      if (!S2(S))
        return c(
          "#VALUE!",
          "Only a LAMBDA formula can be called with arguments.",
        );
      if (O.length !== S.parameters.length)
        return c(
          "#VALUE!",
          `LAMBDA expects ${S.parameters.length} argument(s); received ${O.length}.`,
        );
      if (++o > hc.calls)
        return c("#LIMIT!", "Reusable formula evaluation exceeds 1,024 calls.");
      let R = new Map(S.environment);
      return (
        S.parameters.forEach((I, U) => R.set(I, d(O[U]))),
        _(S.body, S.source, R, S.sheetId, E + 1)
      );
    },
    b = (S, O, E, R, I) => {
      if ((u(1, I), S.kind === "symbol")) return m(S, O, E, R, I);
      if (S.kind === "group") return _(S.body, O, E, R, I + 1);
      if (S.kind === "text") return h(S.text, R, []);
      let U = gp(S);
      if (U === "LET") {
        let v = new Map(E);
        for (let k = 0; k < S.args.length - 1; k += 2) {
          let x = _(S.args[k + 1], O, v, R, I + 1),
            C = mo(Xl(S.args[k]).name);
          (v.set(C, d(x)), i.step?.("binding", `LET ${C}`, x));
        }
        return _(S.args.at(-1), O, v, R, I + 1);
      }
      if (U === "LAMBDA")
        return {
          kind: "formula-closure",
          parameters: S.args.slice(0, -1).map((v) => mo(Xl(v, true).name)),
          body: S.args.at(-1),
          source: O,
          environment: new Map(E),
          sheetId: R,
        };
      let P =
          S.callee.kind === "symbol" && !S.callee.qualifier
            ? Du(S.callee.name)
            : "",
        T =
          S.callee.kind === "symbol" &&
          !S.callee.qualifier &&
          E.has(mo(S.callee.name));
      if (!T && P === "IF" && S.args.length >= 2 && S.args.length <= 3) {
        let v = _(S.args[0], O, E, R, I + 1);
        if (i.isIssue(v)) return v;
        if (!Array.isArray(v)) {
          if (typeof v == "string" && !/^(?:TRUE|FALSE)$/i.test(v))
            return i.issue("#VALUE!");
          let k = typeof v == "string" ? v.toUpperCase() === "TRUE" : !!v;
          return (
            i.step?.(
              "branch",
              `IF: ${k ? "true" : "false"} branch selected`,
              v,
            ),
            S.args[k ? 1 : 2] ? _(S.args[k ? 1 : 2], O, E, R, I + 1) : false
          );
        }
      }
      if (!T && /^(?:IFERROR|IFNA)$/.test(P) && S.args.length === 2) {
        let v = _(S.args[0], O, E, R, I + 1);
        if (!Array.isArray(v)) {
          if (
            i.isIssue(v) &&
            ["#LIMIT!", "#BLOCKED!", "#CYCLE!", "#ERROR!"].includes(v.error)
          )
            return v;
          let k = i.isIssue(v) && (P === "IFERROR" || v.error === "#N/A");
          return (
            i.step?.("branch", `${P}: ${k ? "fallback" : "value"} selected`, v),
            k ? _(S.args[1], O, E, R, I + 1) : v
          );
        }
      }
      if (!T && P && i.supported.has(P)) {
        let v = [],
          k = S.args.map((x) => w(x, O, E, R, I + 1, v));
        return h(`${P}(${k.join(",")})`, R, v);
      }
      let A = b(S.callee, O, E, R, I + 1),
        L = S.args.map((v) => (ql(v).length ? _(v, O, E, R, I + 1) : null)),
        y = g(A, L, I);
      return (
        i.step?.("function", O.slice(S.callee.start, S.callee.end), y),
        y
      );
    },
    w = (S, O, E, R, I, U) =>
      S.map((P) => {
        if (P.kind === "text") {
          let A = /^(\d+(?:\.\d*)?|\.\d+)[eE]([+-]?\d+)$/.exec(P.text);
          return A ? `(${A[1]}*10^(${Number(A[2])}))` : P.text;
        }
        let T = b(P, O, E, R, I);
        return S2(T)
          ? c(
              "#CALC!",
              "A LAMBDA must be called with arguments or saved as a reusable name.",
            )
          : (d(T),
            U.push(
              Array.isArray(T)
                ? T.map((A) => (Array.isArray(A) ? [...A] : A))
                : T,
            ),
            `${Dm}(${U.length - 1})`);
      }).join(""),
    _ = (S, O, E, R, I) => {
      try {
        u(S.length, I);
        let U = ql(S);
        if (U.length === 1 && U[0].kind !== "text") return b(U[0], O, E, R, I);
        let P = [];
        return h(w(S, O, E, R, I, P), R, P);
      } catch (U) {
        if (
          i.isIssue(U) &&
          !["#LIMIT!", "#BLOCKED!", "#CYCLE!", "#ERROR!"].includes(U.error)
        )
          return U;
        throw U;
      }
    };
  try {
    let S = _(f(e), e, new Map(), n, 0);
    return S2(S)
      ? c(
          "#CALC!",
          "A LAMBDA must be called with arguments or saved as a reusable name.",
        )
      : S;
  } catch (S) {
    if (i.isIssue(S)) return S;
    throw S;
  }
}

function Ba(e) {
  throw new Error(e);
}
var $j = (e) =>
  !!e &&
  typeof e == "object" &&
  !Array.isArray(e) &&
  [null, Object.prototype].includes(Object.getPrototypeOf(e));
function Pj(e) {
  return (
    (typeof e != "string" ||
      !e.trim() ||
      e.length > 31 ||
      /[\\/\[\]*?:\u0000-\u001f]/.test(e) ||
      /^'|'$/.test(e)) &&
      Ba(
        "Sheet names must contain 1\u201331 characters and cannot contain control characters, Excel-reserved characters, or begin or end with an apostrophe.",
      ),
    e
  );
}
var Dj = (e) => !e.hidden && !e.veryHidden;
function Lj(e) {
  return (
    (typeof e != "string" || !/^#[a-f\d]{6}$/i.test(e)) &&
      Ba("Tab colors must use a six-digit hexadecimal color."),
    e.toLowerCase()
  );
}
function E5(e, t) {
  (Pj(t.name),
    e.veryHidden !== undefined &&
      (typeof e.veryHidden != "boolean" &&
        Ba("The very-hidden flag must be boolean."),
      e.veryHidden &&
        e.hidden === false &&
        Ba("A very-hidden worksheet must also be hidden."),
      (t.veryHidden = e.veryHidden),
      e.veryHidden && (t.hidden = true)),
    e.tabColor !== undefined && (t.tabColor = Lj(e.tabColor)));
}
function S5(e, t) {
  let n = e;
  if (n !== undefined) {
    (!$j(n) ||
      n.enabled !== true ||
      Object.keys(n).some(
        (o) => !["enabled", "passwordHash", "salt"].includes(o),
      )) &&
      Ba("Invalid workbook structure protection.");
    let i = n.passwordHash !== undefined,
      r = n.salt !== undefined;
    ((i !== r ||
      (i &&
        (typeof n.passwordHash != "string" ||
          !/^[a-f\d]{64}$/i.test(n.passwordHash) ||
          typeof n.salt != "string" ||
          n.salt.length < 16 ||
          n.salt.length > 100))) &&
      Ba(
        "Workbook structure protection requires a valid salted password hash.",
      ),
      (t.structureProtection = {
        enabled: true,
        ...(i
          ? {
              passwordHash: String(n.passwordHash).toLowerCase(),
              salt: String(n.salt),
            }
          : {}),
      }));
  }
  t.sheets.some(Dj) || Ba("At least one worksheet must remain visible.");
}
var U2 = String.raw`(?:'(?:[^']|'')+'|[^\s+\-*/^&=<>(),;:'"\[\]!]+)`,
  qQ = new RegExp(`^(${U2})!`),
  XQ = new RegExp(`^(${U2}):(${U2})!`);
var Fj = ["any", "text", "number", "boolean", "date"];
function Fm(e) {
  if (!Fj.includes(e)) throw Error("Choose a supported table column type.");
  return e;
}
function zm(e, t) {
  if (e === undefined) return;
  let n = () => {
      throw Error(
        "Invalid or over-capacity table source evidence. Keep the native backup.",
      );
    },
    i = (a) => !!a && typeof a == "object" && !Array.isArray(a);
  (!Array.isArray(e) || e.length > 100 || JSON.stringify(e).length > 5e6) &&
    n();
  let r = new Set(),
    o = 0;
  t.tableEvidence = e.map((a) => {
    (!i(a) ||
      Object.keys(a).some(
        (u) =>
          ![
            "id",
            "kind",
            "sourceRevision",
            "sheetId",
            "sheetName",
            "tableId",
            "tableName",
            "columnId",
            "dataType",
            "changes",
          ].includes(u),
      )) &&
      n();
    let s = a;
    for (let u of ["id", "sheetId", "sheetName", "tableId", "tableName"])
      (typeof s[u] != "string" || !s[u] || s[u].length > 512) && n();
    ((r.has(s.id) ||
      !Number.isSafeInteger(s.sourceRevision) ||
      Number(s.sourceRevision) < 0 ||
      !["headers", "columnType"].includes(String(s.kind)) ||
      !Array.isArray(s.changes) ||
      s.changes.length > 500 ||
      (o += s.changes.length) > 1e4) &&
      n(),
      s.kind === "columnType"
        ? ((typeof s.columnId != "string" ||
            !s.columnId ||
            s.columnId.length > 100) &&
            n(),
          Fm(s.dataType))
        : (s.columnId !== undefined || s.dataType !== undefined) && n(),
      r.add(s.id));
    let l = new Set(),
      c = s.changes.map((u) => {
        (!i(u) ||
          Object.keys(u).some(
            (h) => !["address", "before", "after"].includes(h),
          ) ||
          typeof u.address != "string") &&
          n();
        let d = u,
          f = d.address,
          p = tt(f);
        return (
          (`${Ue(p.col)}${p.row + 1}` !== f || l.has(f)) && n(),
          l.add(f),
          {
            address: f,
            ...(d.before !== undefined ? { before: Oi(d.before) } : {}),
            after: Oi(d.after),
          }
        );
      });
    return { ...s, changes: c };
  });
}
var Fs = (e) => {
    throw Error(e);
  },
  zj = (e) => typeof e == "number" && Number.isFinite(e);
var C5 = (e) => !!e && typeof e == "object" && !Array.isArray(e);
function T5(e) {
  C5(e) || Fs("Invalid pivot refresh evidence.");
  let t = e;
  (typeof t.at != "string" ||
    !/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}Z$/.test(t.at) ||
    !Number.isFinite(Date.parse(t.at)) ||
    new Date(t.at).toISOString() !== t.at ||
    typeof t.sourceSheetName != "string" ||
    !t.sourceSheetName ||
    t.sourceSheetName.length > 31 ||
    typeof t.sourceRange != "string" ||
    !/^[A-Z]+\d+:[A-Z]+\d+$/.test(t.sourceRange)) &&
    Fs("Invalid pivot refresh time or source label.");
  for (let i of ["sourceFingerprint", "outputFingerprint"])
    (typeof t[i] != "string" || !/^[a-f0-9]{64}$/.test(t[i])) &&
      Fs("Invalid pivot content fingerprint.");
  for (let i of ["sourceRows", "includedRows", "excludedRows"])
    (!Number.isSafeInteger(t[i]) || Number(t[i]) < 0 || Number(t[i]) > 1e5) &&
      Fs("Invalid pivot population counts.");
  (Number(t.sourceRows) !== Number(t.includedRows) + Number(t.excludedRows) ||
    !Array.isArray(t.measures) ||
    !t.measures.length ||
    t.measures.length > 100) &&
    Fs("Pivot population evidence does not reconcile.");
  let n = t.measures.map((i) => {
    (!C5(i) ||
      typeof i.name != "string" ||
      !i.name ||
      i.name.length > 512 ||
      typeof i.field != "string" ||
      !i.field ||
      i.field.length > 255 ||
      !["sum", "count", "average", "min", "max", "distinctCount"].includes(
        String(i.operation),
      ) ||
      !(
        i.result === null ||
        typeof i.result == "boolean" ||
        zj(i.result) ||
        (typeof i.result == "string" && i.result.length <= 32767)
      )) &&
      Fs("Invalid pivot measure evidence.");
    let r = i;
    for (let o of ["numericCount", "nonblankCount"])
      (!Number.isSafeInteger(r[o]) ||
        Number(r[o]) < 0 ||
        Number(r[o]) > Number(t.includedRows)) &&
        Fs("Invalid pivot measure count.");
    return (
      Number(r.numericCount) > Number(r.nonblankCount) &&
        Fs("Invalid pivot numeric count."),
      {
        name: r.name,
        field: r.field,
        operation: r.operation,
        result: r.result,
        numericCount: r.numericCount,
        nonblankCount: r.nonblankCount,
      }
    );
  });
  return {
    at: t.at,
    sourceSheetName: t.sourceSheetName,
    sourceRange: t.sourceRange,
    sourceFingerprint: t.sourceFingerprint,
    outputFingerprint: t.outputFingerprint,
    sourceRows: t.sourceRows,
    includedRows: t.includedRows,
    excludedRows: t.excludedRows,
    measures: n,
  };
}
var I5 = ["value", "percentGrand", "percentRow", "percentColumn"],
  os = (e) => {
    throw Error(e);
  },
  Bm = (e) => !!e && typeof e == "object" && !Array.isArray(e),
  F2 = (e) =>
    typeof e == "string" && e.trim() && e.length <= 255
      ? e
      : os("Choose a valid pivot field."),
  A5 = (e) =>
    typeof e == "string" &&
    /^\d{4}-\d{2}-\d{2}$/.test(e) &&
    !e.startsWith("0000") &&
    Number.isFinite(Date.parse(e)) &&
    new Date(e).toISOString().slice(0, 10) === e,
  jm = (e, t) =>
    Array.isArray(e) && e.length <= t
      ? e
      : os(`Pivot controls require a list of at most ${t} entries.`),
  R5 = (e) =>
    JSON.stringify([
      typeof e,
      typeof e == "string" ? e.toLocaleLowerCase("en-US") : e,
    ]);
function N5(e) {
  let t = {};
  (e.slicers !== undefined &&
    (t.slicers = jm(e.slicers, 20).map((n) => {
      if (!Bm(n)) return os("Invalid pivot slicer.");
      let i = jm(n.selected, 500).map(
        (r) => (
          r !== null &&
            typeof r != "boolean" &&
            !(typeof r == "number" && Number.isFinite(r)) &&
            !(typeof r == "string" && r.length <= 32767) &&
            os("Invalid slicer item."),
          r
        ),
      );
      return (
        new Set(i.map(R5)).size !== i.length &&
          os("Slicer selections must be unique."),
        { field: F2(n.field), selected: i }
      );
    })),
    e.timelines !== undefined &&
      (t.timelines = jm(e.timelines, 10).map((n) =>
        !Bm(n) ||
        !A5(n.from) ||
        !A5(n.to) ||
        n.from > n.to ||
        !["year", "quarter", "month"].includes(String(n.unit)) ||
        !["iso", "1900", "1904"].includes(String(n.dateSystem))
          ? os(
              "Timeline requires ordered YYYY-MM-DD bounds, a period and an explicit date system.",
            )
          : {
              field: F2(n.field),
              from: n.from,
              to: n.to,
              unit: n.unit,
              dateSystem: n.dateSystem,
            },
      )),
    e.sorts !== undefined &&
      (t.sorts = jm(e.sorts, 100).map((n) =>
        !Bm(n) || !["ascending", "descending"].includes(String(n.direction))
          ? os("Pivot field sorting requires ascending or descending order.")
          : { field: F2(n.field), direction: n.direction },
      )));
  for (let n of [t.slicers, t.timelines, t.sorts])
    n &&
      new Set(n.map((i) => i.field)).size !== n.length &&
      os("Choose only one control of each kind per field.");
  if (e.chart !== undefined) {
    let n = e.chart;
    if (
      !Bm(n) ||
      !["column", "bar", "line"].includes(String(n.kind)) ||
      !Number.isInteger(n.valueIndex) ||
      Number(n.valueIndex) < 0 ||
      Number(n.valueIndex) >= Number(e.values?.length)
    )
      return os(
        "PivotChart requires a supported chart and an existing value field.",
      );
    t.chart = { kind: n.kind, valueIndex: n.valueIndex };
  }
  return t;
}
var O5 = (e) =>
    !!e &&
    typeof e == "object" &&
    !Array.isArray(e) &&
    (Object.getPrototypeOf(e) === Object.prototype ||
      Object.getPrototypeOf(e) === null),
  bc = (e, t) => typeof e == "string" && e.length > 0 && e.length <= t;
function M5(e, t) {
  let n = e.sheets;
  if (
    e.rowOrigins === undefined &&
    !n.some((o) => o.rowOriginIds !== undefined)
  )
    return;
  if (!Array.isArray(e.rowOrigins) || e.rowOrigins.length > Ne.maxCells)
    throw Error(
      "Row provenance requires a registry of at most 100,000 source rows.",
    );
  let i = new Set();
  t.rowOrigins = e.rowOrigins.map((o) => {
    if (
      !O5(o) ||
      !bc(o.id, 100) ||
      i.has(o.id) ||
      !bc(o.sourceId, 100) ||
      !bc(o.sourceName, 255) ||
      !bc(o.sourceSheet, 31) ||
      !["file", "snapshot"].includes(String(o.kind)) ||
      !Number.isInteger(o.originalRow) ||
      Number(o.originalRow) < 1 ||
      Number(o.originalRow) > Ne.maxRows
    )
      throw Error("Invalid, duplicated or out-of-bounds original source row.");
    return (
      i.add(o.id),
      {
        id: o.id,
        sourceId: o.sourceId,
        sourceName: o.sourceName,
        sourceSheet: o.sourceSheet,
        originalRow: o.originalRow,
        kind: o.kind,
      }
    );
  });
  let r = 0;
  for (let o = 0; o < n.length; o++) {
    let a = n[o].rowOriginIds;
    if (a === undefined) continue;
    if (!O5(a))
      throw Error("Source row associations must be a plain index map.");
    let s = {};
    for (let [l, c] of Object.entries(a)) {
      if (
        !/^(0|[1-9]\d*)$/.test(l) ||
        Number(l) >= Number(n[o].rowCount) ||
        typeof c != "string" ||
        !i.has(c) ||
        ++r > Ne.maxCells
      )
        throw Error("Invalid or over-capacity source row association.");
      s[l] = c;
    }
    t.sheets[o].rowOriginIds = s;
  }
}
function Wm(e, t) {
  if (
    !bc(t.id, 100) ||
    !bc(t.name, 255) ||
    !["file", "snapshot"].includes(t.kind)
  )
    throw Error("Provide a valid source identity, name and origin kind.");
  let n = [...(e.rowOrigins ?? [])],
    i = e.sheets.map((r) => {
      let o = { ...r.rowOriginIds },
        a = new Set(
          Object.keys(r.cells).map((s) => Number(/\d+$/.exec(s)[0]) - 1),
        );
      for (let s of a)
        if (!o[s]) {
          if (n.length >= Ne.maxCells)
            throw Error(
              "Row provenance has reached its 100,000 source-row budget.",
            );
          let l = crypto.randomUUID();
          ((o[s] = l),
            n.push({
              id: l,
              sourceId: t.id,
              sourceName: t.name.slice(0, 255),
              sourceSheet: r.name,
              originalRow: s + 1,
              kind: t.kind,
            }));
        }
      return { ...r, rowOriginIds: o };
    });
  return { ...e, sheets: i, rowOrigins: n };
}
var B2 = ["sum", "count", "average", "min", "max"],
  jj = [...B2, "distinctCount"],
  $5 = () => crypto.randomUUID();
function ut(e) {
  throw new Error(e);
}
var Sn = (e) =>
    !!e &&
    typeof e == "object" &&
    !Array.isArray(e) &&
    (Object.getPrototypeOf(e) === Object.prototype ||
      Object.getPrototypeOf(e) === null),
  oo = (e, t, n = 255) =>
    typeof e == "string" && e.trim() && e.length <= n
      ? e
      : ut(`${t} must contain 1\u2013${n} characters.`),
  yc = (e) => oo(e, "Metadata ID", 100),
  Vm = (e) => {
    let t = Qe(e);
    return `${Ue(t.firstCol)}${t.firstRow + 1}:${Ue(t.lastCol)}${t.lastRow + 1}`;
  },
  qm = (e, t) => `${Ue(t)}${e + 1}`,
  Wj = (e, t, n) =>
    t >= e.firstRow && t <= e.lastRow && n >= e.firstCol && n <= e.lastCol,
  Vj = (e, t) =>
    e.firstRow <= t.lastRow &&
    e.lastRow >= t.firstRow &&
    e.firstCol <= t.lastCol &&
    e.lastCol >= t.firstCol,
  D5 = (e) => (e.lastRow - e.firstRow + 1) * (e.lastCol - e.firstCol + 1);
var P5 = {
  blue: "TableStyleMedium2",
  green: "TableStyleMedium4",
  gray: "TableStyleMedium9",
};
function Xm(e, t) {
  let n = Qe(oo(e, "Source range", 50));
  return (
    D5(n) > Ne.maxCells &&
      ut(
        "Table, query and pivot ranges are limited to 100,000 cells including blank gaps. Select the used data range.",
      ),
    t &&
      (n.lastRow >= t.rowCount || n.lastCol >= t.colCount) &&
      ut("Range exceeds worksheet bounds."),
    n
  );
}
function Bn(e) {
  let t = oo(e, "Column name");
  return (
    /[\[\]\r\n]/.test(t) &&
      ut("Column names cannot contain brackets or line breaks."),
    t
  );
}
function L5(e) {
  let t = oo(e, "Table name");
  return (
    (!/^[A-Za-z_][A-Za-z0-9_.]*$/.test(t) ||
      /^[RC]$/i.test(t) ||
      /^[A-Z]{1,3}[1-9]\d*$/i.test(t) ||
      /^R\d+C\d+$/i.test(t)) &&
      ut(
        "Use a table name beginning with a letter or underscore, without spaces or a cell-address name.",
      ),
    t
  );
}
function z2(e, t) {
  (!Array.isArray(e) || e.length > Ne.maxCols) &&
    ut(`${t} must be a bounded list of column names.`);
  let n = e.map(Bn);
  return (
    new Set(n.map((i) => i.toLowerCase())).size !== n.length &&
      ut(`${t} contains duplicate fields.`),
    n
  );
}
function U5(e) {
  Sn(e) || ut("A source worksheet and range are required.");
  let t = {
    sheetId: yc(e.sheetId),
    range: Vm(oo(e.range, "Source range", 50)),
    ...(e.tableId !== undefined ? { tableId: yc(e.tableId) } : {}),
  };
  return (Xm(t.range), t);
}
function F5(e) {
  return e === undefined
    ? []
    : ((!Array.isArray(e) || e.length > 100) &&
        ut("Use at most 100 source filters."),
      e.map((t) => {
        (!Sn(t) ||
          ![
            "equals",
            "notEquals",
            "contains",
            "gt",
            "gte",
            "lt",
            "lte",
            "blank",
            "notBlank",
          ].includes(String(t.operator))) &&
          ut("Invalid source filter.");
        let n = t.value;
        return (
          n != null &&
            !(typeof n == "string" && n.length <= 32767) &&
            typeof n != "boolean" &&
            !(typeof n == "number" && Number.isFinite(n)) &&
            ut(
              "Filter values must be finite numbers, text, booleans or blank.",
            ),
          {
            field: Bn(t.field),
            operator: t.operator,
            ...(n !== undefined ? { value: n } : {}),
          }
        );
      }));
}
function z5(e) {
  if (e === undefined) return;
  Sn(e) || ut("Invalid output destination.");
  let t = e.address === undefined ? "A1" : oo(e.address, "Output address", 20),
    n = tt(t);
  return {
    address: qm(n.row, n.col),
    ...(e.sheetId !== undefined ? { sheetId: yc(e.sheetId) } : {}),
    ...(e.sheetName !== undefined
      ? { sheetName: oo(e.sheetName, "Output sheet name", 31) }
      : {}),
  };
}
function qj(e) {
  (Sn(e) || ut("Invalid pivot definition."),
    (!Array.isArray(e.values) || !e.values.length || e.values.length > 100) &&
      ut("Choose between 1 and 100 pivot value fields."));
  let t = e.values.map(
    (r) => (
      (!Sn(r) || !jj.includes(r.operation)) &&
        ut(
          "Invalid pivot aggregation. Choose sum, count, average, min, max or distinctCount.",
        ),
      r.showAs !== undefined &&
        !I5.includes(r.showAs) &&
        ut("Choose a supported Show values as setting."),
      {
        field: Bn(r.field),
        operation: r.operation,
        ...(r.name !== undefined ? { name: Bn(r.name) } : {}),
        ...(r.showAs !== undefined ? { showAs: r.showAs } : {}),
      }
    ),
  );
  (e.grandTotals !== undefined &&
    typeof e.grandTotals != "boolean" &&
    ut("Grand totals must be boolean."),
    e.subtotals !== undefined &&
      typeof e.subtotals != "boolean" &&
      ut("Subtotals must be boolean."),
    e.dateGroups !== undefined &&
      (!Array.isArray(e.dateGroups) || e.dateGroups.length > 100) &&
      ut("Use at most 100 date groups."));
  let n = (e.dateGroups ?? []).map(
    (r) => (
      (!Sn(r) ||
        !["year", "quarter", "month"].includes(String(r.unit)) ||
        !["iso", "1900", "1904"].includes(String(r.dateSystem))) &&
        ut(
          "Date grouping requires a year, quarter or month and an explicit ISO/1900/1904 date system.",
        ),
      { field: Bn(r.field), unit: r.unit, dateSystem: r.dateSystem }
    ),
  );
  (new Set(n.map((r) => r.field)).size !== n.length &&
    ut("Choose one date grouping per source field."),
    e.calculatedFields !== undefined &&
      (!Array.isArray(e.calculatedFields) || e.calculatedFields.length > 100) &&
      ut("Use at most 100 calculated fields."));
  let i = (e.calculatedFields ?? []).map(
    (r) => (
      (!Sn(r) ||
        !["add", "subtract", "multiply", "divide"].includes(
          String(r.operation),
        ) ||
        (r.otherField === undefined) == (r.constant === undefined) ||
        (r.constant !== undefined &&
          (typeof r.constant != "number" || !Number.isFinite(r.constant)))) &&
        ut(
          "A calculated field needs add/subtract/multiply/divide and exactly one numeric constant or other source field.",
        ),
      {
        name: Bn(r.name),
        field: Bn(r.field),
        operation: r.operation,
        ...(r.otherField !== undefined
          ? { otherField: Bn(r.otherField) }
          : { constant: r.constant }),
      }
    ),
  );
  return (
    new Set(i.map((r) => r.name.toLowerCase())).size !== i.length &&
      ut("Calculated field names must be unique."),
    {
      name: oo(e.name, "Pivot name", 100),
      source: U5(e.source),
      rowFields: z2(e.rowFields, "Row fields"),
      columnFields: z2(e.columnFields, "Column fields"),
      values: t,
      filters: F5(e.filters),
      grandTotals: e.grandTotals !== false,
      ...(e.subtotals !== undefined ? { subtotals: e.subtotals } : {}),
      ...(e.dateGroups !== undefined ? { dateGroups: n } : {}),
      ...(e.calculatedFields !== undefined ? { calculatedFields: i } : {}),
      ...N5(e),
      destination: z5(e.destination),
    }
  );
}
function Xj(e) {
  Sn(e) || ut("Invalid query definition.");
  let t = z2(e.fields, "Selected fields");
  t.length || ut("Choose at least one output field.");
  let n = {};
  if (e.rename !== undefined) {
    Sn(e.rename) || ut("Rename fields must be a name map.");
    for (let [a, s] of Object.entries(e.rename))
      (t.includes(a) || ut("Only selected fields can be renamed."),
        Object.defineProperty(n, a, {
          value: Bn(s),
          enumerable: true,
          writable: true,
          configurable: true,
        }));
  }
  let i;
  (e.sort !== undefined &&
    ((!Sn(e.sort) ||
      !["ascending", "descending"].includes(String(e.sort.direction))) &&
      ut("Invalid query sort."),
    (i = { field: Bn(e.sort.field), direction: e.sort.direction })),
    e.derived !== undefined &&
      (!Array.isArray(e.derived) || e.derived.length > 100) &&
      ut("Use at most 100 derived fields."));
  let r = (e.derived ?? []).map(
      (a) => (
        (!Sn(a) ||
          ![
            "add",
            "subtract",
            "multiply",
            "divide",
            "concat",
            "upper",
            "lower",
            "trim",
          ].includes(String(a.operation))) &&
          ut("Invalid derived operation."),
        a.constant !== undefined &&
          !(typeof a.constant == "string" && a.constant.length <= 32767) &&
          !(typeof a.constant == "number" && Number.isFinite(a.constant)) &&
          ut("A derived constant must be bounded text or a finite number."),
        {
          name: Bn(a.name),
          field: Bn(a.field),
          operation: a.operation,
          ...(a.otherField !== undefined
            ? { otherField: Bn(a.otherField) }
            : {}),
          ...(a.constant !== undefined ? { constant: a.constant } : {}),
        }
      ),
    ),
    o = [
      ...t.map((a) => (Object.hasOwn(n, a) ? n[a] : a)),
      ...r.map((a) => a.name),
    ];
  return (
    new Set(o.map((a) => a.toLowerCase())).size !== o.length &&
      ut("Output column names must be unique."),
    {
      name: oo(e.name, "Query name", 100),
      source: U5(e.source),
      fields: t,
      rename: n,
      filters: F5(e.filters),
      ...(i ? { sort: i } : {}),
      derived: r,
      destination: z5(e.destination),
    }
  );
}
function j2(e, t) {
  if (Sn(e)) {
    if (e.tables !== undefined) {
      (!Array.isArray(e.tables) || e.tables.length > Ne.maxCells) &&
        ut("Tables must be a bounded list.");
      let n = new Set(),
        i = new Set();
      t.tables = e.tables.map((a) => {
        (!Sn(a) ||
          !Array.isArray(a.columns) ||
          typeof a.showTotals != "boolean" ||
          !["blue", "green", "gray"].includes(String(a.style)) ||
          !Sn(a.totals)) &&
          ut("Invalid table metadata.");
        let s = yc(a.id),
          l = L5(a.name),
          c = Vm(oo(a.range, "Table range", 50)),
          u = Xm(c, t);
        (t.merges?.some((p) => Vj(u, Qe(p))) &&
          ut("A table cannot overlap merged cells. Unmerge the source first."),
          (n.has(s) || i.has(l.toLowerCase())) &&
            ut("Table IDs and names must be unique."),
          n.add(s),
          i.add(l.toLowerCase()),
          (a.columns.length !== u.lastCol - u.firstCol + 1 ||
            (a.showTotals && u.lastRow <= u.firstRow)) &&
            ut("Table columns or totals do not fit the table range."));
        let d = a.columns.map(
          (p) => (
            Sn(p) || ut("Invalid table column."),
            {
              id: yc(p.id),
              name: Bn(p.name),
              ...(p.dataType !== undefined ? { dataType: Fm(p.dataType) } : {}),
              ...(p.calculatedFormula !== undefined
                ? { calculatedFormula: Zj(p.calculatedFormula, true) }
                : {}),
            }
          ),
        );
        (new Set(d.map((p) => p.id)).size !== d.length ||
          new Set(d.map((p) => p.name.toLowerCase())).size !== d.length) &&
          ut("Table column IDs and names must be unique.");
        let f = {};
        for (let [p, h] of Object.entries(a.totals))
          ((!d.some((m) => m.name === p) || !B2.includes(h)) &&
            ut("Invalid table total field or operation."),
            Object.defineProperty(f, p, {
              value: h,
              enumerable: true,
              writable: true,
              configurable: true,
            }));
        return {
          id: s,
          name: l,
          range: c,
          columns: d,
          style: a.style,
          showTotals: a.showTotals,
          totals: f,
        };
      });
      let r = new Set(),
        o = 0;
      for (let a of t.tables) {
        let s = Qe(a.range);
        (o += D5(s)) > Ne.maxCells &&
          ut("The total table footprint exceeds 100,000 cells.");
        for (let l = s.firstRow; l <= s.lastRow; l++)
          for (let c = s.firstCol; c <= s.lastCol; c++) {
            let u = qm(l, c);
            (r.has(u) && ut("Tables cannot overlap."), r.add(u));
          }
      }
    }
    for (let n of ["pivots", "queries"]) {
      if (e[n] === undefined) continue;
      (!Array.isArray(e[n]) || e[n].length > Ne.maxCells) &&
        ut("Analysis definitions must be a bounded list.");
      let i = new Set(),
        r = e[n].map((o) => {
          Sn(o) || ut("Invalid analysis metadata.");
          let a = n === "pivots" ? qj(o) : Xj(o),
            s = yc(o.id),
            l = Vm(oo(o.outputRange, "Output range", 50));
          (Xm(l, t),
            (i.has(s) ||
              !Number.isSafeInteger(o.sourceRevision) ||
              Number(o.sourceRevision) < 0) &&
              ut("Invalid analysis identity or source revision."),
            i.add(s));
          let { destination, ...u } = a;
          return {
            ...u,
            id: s,
            outputRange: l,
            sourceRevision: Number(o.sourceRevision),
            ...(n === "pivots" && o.refresh !== undefined
              ? { refresh: T5(o.refresh) }
              : {}),
          };
        });
      n === "pivots" ? (t.pivots = r) : (t.queries = r);
    }
  }
}
function B5(e) {
  let t = new Set(),
    n = new Set();
  for (let i of e.sheets)
    for (let r of i.tables ?? [])
      ((t.has(r.name.toLowerCase()) || n.has(r.id)) &&
        ut("Table names and IDs must be unique across the workbook."),
        t.add(r.name.toLowerCase()),
        n.add(r.id));
}
function Hj(e) {
  let t = Qe(e.range);
  return {
    ...t,
    firstRow: t.firstRow + 1,
    lastRow: t.lastRow - (e.showTotals ? 1 : 0),
  };
}
function Hm(e, t, n, i) {
  let r = /^([A-Za-z_][A-Za-z0-9_.]*)?(\[.+\])$/.exec(n);
  if (!r) return null;
  let o;
  if (r[1])
    for (let h of e.sheets) {
      let m = h.tables?.find(
        (g) => g.name.toLowerCase() === r[1].toLowerCase(),
      );
      if (m) {
        o = { sheet: h, table: m };
        break;
      }
    }
  else if (i) {
    let h = e.sheets.find((b) => b.id === t),
      m = tt(i),
      g = h?.tables?.find((b) => Wj(Qe(b.range), m.row, m.col));
    h && g && (o = { sheet: h, table: g });
  }
  if (!o) return null;
  let { sheet, table } = o,
    l = Qe(table.range),
    c = Hj(table),
    u = r[2]
      .replace(/^\[|\]$/g, "")
      .replace(/[\[\]]/g, "")
      .split(/[:,]/)
      .map((h) => h.trim()),
    d = u.find((h) => h.startsWith("#"))?.toLowerCase(),
    f = u.some((h) => h.startsWith("@")) || d === "#this row";
  if (
    u.some(
      (h) =>
        h.startsWith("#") &&
        !["#all", "#data", "#headers", "#totals", "#this row"].includes(
          h.toLowerCase(),
        ),
    )
  )
    return null;
  if (
    (d === "#all" && Object.assign(c, l),
    d === "#headers" && Object.assign(c, l, { lastRow: l.firstRow }),
    d === "#totals")
  ) {
    if (!table.showTotals) return null;
    Object.assign(c, l, { firstRow: l.lastRow });
  }
  if (f) {
    if (!i || sheet.id !== t) return null;
    let h = tt(i);
    if (h.row < c.firstRow || h.row > c.lastRow) return null;
    ((c.firstRow = h.row), (c.lastRow = h.row));
  }
  let p = u.filter((h) => !h.startsWith("#")).map((h) => h.replace(/^@/, ""));
  if (p.length) {
    if (p.length > 2) return null;
    let h = p.map((m) =>
      table.columns.findIndex((g) => g.name.toLowerCase() === m.toLowerCase()),
    );
    if (h.some((m) => m < 0)) return null;
    ((c.firstCol = l.firstCol + Math.min(...h)),
      (c.lastCol = l.firstCol + Math.max(...h)));
  }
  return c.firstRow > c.lastRow
    ? null
    : {
        sheetId: sheet.id,
        range: `${qm(c.firstRow, c.firstCol)}:${qm(c.lastRow, c.lastCol)}`,
      };
}
function j5(e, t) {
  let n = [],
    i = [],
    r = e.model.tables ?? [];
  for (let o of r) {
    if (o.headerRow === false) {
      n.push(
        `Table ${o.name} has no header row; its cells were retained without table metadata.`,
      );
      continue;
    }
    try {
      let a = Vm(o.tableRef);
      Xm(a, t);
      let s =
          Object.entries(P5).find(([, c]) => c === o.style?.theme)?.[0] ??
          "blue",
        l = {};
      for (let c of o.columns) {
        let u = c.totalsRowFunction?.toLowerCase();
        if (B2.includes(u))
          Object.defineProperty(l, c.name, {
            value: u,
            enumerable: true,
            writable: true,
            configurable: true,
          });
        else if (u === "custom") {
          let d = c.totalsRowFormula ?? "",
            f = /^=?\s*(SUM|COUNTA|AVERAGE|MIN|MAX)\(/i
              .exec(d)?.[1]
              .toLowerCase();
          f &&
            Object.defineProperty(l, c.name, {
              value: f === "counta" ? "count" : f,
              enumerable: true,
              writable: true,
              configurable: true,
            });
        }
      }
      (i.push({
        id: $5(),
        name: L5(o.name),
        range: a,
        columns: o.columns.map((c) => ({ id: $5(), name: Bn(c.name) })),
        style: s,
        showTotals: !!o.totalsRow,
        totals: l,
      }),
        Object.values(P5).includes(o.style?.theme) ||
          n.push(
            `Table ${o.name}: source cells were preserved; unsupported table theme maps to the blue table style for future edits.`,
          ));
    } catch (a) {
      n.push(
        `Table ${o.name}: ${a.message} Cells remain available without this table definition.`,
      );
    }
  }
  return (i.length && ((t.tables = i), j2(t, t)), n);
}
var Gj = { changedCells: 500, appendedRows: 100, formulaLength: 4096 };
function Zj(e, t = false) {
  (typeof e != "string" ||
    e.length > Gj.formulaLength ||
    e !== e.trim() ||
    !e.startsWith("=")) &&
    ut(
      "Calculated column formulas must begin with = and contain at most 4,096 characters, without outer whitespace.",
    );
  try {
    Cd(e);
  } catch (n) {
    ut(
      `Calculated column formula: ${n.message.replace(/^XLSX defined names: /, "").replace(/ Import\/export stopped.*$/, "")}`,
    );
  }
  return (
    !t &&
      /#REF!/.test(Co(e.replace(/"(?:[^"]|"")*"/g, '""'))) &&
      ut(
        "Repair broken #REF! references before saving a calculated column rule.",
      ),
    e
  );
}
var W5 = "\\$?[A-Za-z]{1,3}\\$?[1-9]\\d{0,6}",
  Kj = new RegExp(`^(${W5}(?::${W5})?)(?![A-Za-z0-9_.$:#\\[])`);
function Zm(e, t = 5e3) {
  let n = (p) => {
      let h = /^\$?([A-Za-z]{1,3})\$?([1-9]\d{0,6})$/.exec(p);
      if (!h)
        throw Error("External links require bounded A1 cells or rectangles.");
      let m = 0;
      for (let b of h[1].toUpperCase()) m = m * 26 + b.charCodeAt(0) - 64;
      let g = Number(h[2]);
      if (m > 16384 || g > 1048576)
        throw Error("External reference exceeds worksheet bounds.");
      return { row: g, col: m };
    },
    [i, r = i, o] = e.split(":");
  if (o) throw Error("Invalid external range.");
  let a = n(i),
    s = n(r),
    l = [],
    c = Math.min(a.row, s.row),
    u = Math.max(a.row, s.row),
    d = Math.min(a.col, s.col),
    f = Math.max(a.col, s.col);
  if ((u - c + 1) * (f - d + 1) > t)
    throw Error(
      `External link ranges are limited to ${t.toLocaleString()} cells.`,
    );
  for (let p = c; p <= u; p++)
    for (let h = d; h <= f; h++) {
      let m = h,
        g = "";
      for (; m;)
        ((g = String.fromCharCode(65 + ((m - 1) % 26)) + g),
          (m = Math.floor((m - 1) / 26)));
      l.push(g + p);
    }
  return l;
}
function Jj(e) {
  let t = [];
  for (let n = 0; n < e.length;) {
    if (e[n] === '"') {
      for (n++; n < e.length;)
        if (e[n++] === '"')
          if (e[n] === '"') n++;
          else break;
      continue;
    }
    let i = n,
      r = /^'((?:[^']|'')*)'\s*!/.exec(e.slice(n)),
      o = /^\[([^\[\]\r\n]+)\]([A-Za-z_][A-Za-z0-9_.]*)!/.exec(e.slice(n)),
      a,
      s = 0;
    if (
      (r
        ? ((a = r[1].replace(/''/g, "'")), (s = r[0].length))
        : o && ((a = `[${o[1]}]${o[2]}`), (s = o[0].length)),
      a)
    ) {
      let l = /^(.*\[[^\[\]\r\n]+\])([^\[\]:\r\n]+)$/.exec(a),
        c = Kj.exec(e.slice(n + s));
      if (l && c && (!n || !/[A-Za-z0-9_.\]]/.test(e[n - 1])))
        try {
          (Zm(c[1]),
            (n += s + c[0].length),
            t.push({
              start: i,
              end: n,
              token: e.slice(i, n),
              target: l[1],
              sheet: l[2],
              range: c[1].toUpperCase(),
            }));
          continue;
        } catch {}
      n += s;
      continue;
    }
    if (e[n] === "[") {
      let l = 0;
      do (e[n] === "[" && l++, e[n] === "]" && l--, n++);
      while (n < e.length && l > 0);
    } else if (e[n] === "'") {
      for (n++; n < e.length;)
        if (e[n++] === "'")
          if (e[n] === "'") n++;
          else break;
    } else n++;
  }
  return t;
}
function V5(e, t) {
  let n = Jj(e);
  for (let i of n.reverse()) {
    let r = t(i);
    r !== undefined && (e = e.slice(0, i.start) + r + e.slice(i.end));
  }
  return e;
}
var V2 = Object.freeze({
    sources: 8,
    sheets: 32,
    cells: 5e3,
    formulas: 500,
    metadataCharacters: 2e6,
  }),
  ja = (e) =>
    X.string()
      .min(1)
      .max(e)
      .refine((t) => !/[\u0000-\u001f\u007f]/.test(t)),
  X5 = ja(2048).refine(
    (e) => /^[^\[\]]*\[[^\[\]]+\]$/.test(e),
    "Use an exact workbook qualifier such as [Source.xlsx].",
  ),
  H5 = X.string().regex(/^[a-f0-9]{64}$/),
  Yj = X.discriminatedUnion("type", [
    X.strictObject({ type: X.literal("number"), value: X.number().finite() }),
    X.strictObject({ type: X.literal("text"), value: X.string().max(32767) }),
    X.strictObject({ type: X.literal("boolean"), value: X.boolean() }),
    X.strictObject({ type: X.literal("blank"), value: X.null() }),
    X.strictObject({
      type: X.literal("error"),
      value: X.string()
        .regex(/^#[A-Z0-9/?!_.]+$/)
        .max(30),
    }),
  ]),
  Qj = X.strictObject({
    target: X5,
    sourceId: ja(100),
    sourceName: ja(200),
    sourceRevision: X.number().int().nonnegative(),
    sourceFingerprint: H5,
    resolvedAt: X.iso.datetime(),
    sheets: X.array(
      X.strictObject({
        id: ja(100),
        name: ja(31),
        cells: X.record(X.string().regex(/^[A-Z]{1,3}[1-9]\d{0,6}$/), Yj),
      }),
    )
      .min(1)
      .max(32),
  }),
  eW = X.strictObject({
    schemaVersion: X.literal(1),
    sources: X.array(Qj).max(8),
    packageTargets: X.array(
      X.strictObject({
        key: X5,
        target: ja(4096),
        part: ja(1024),
        sourceSha256: H5,
      }),
    )
      .max(100)
      .optional(),
  }),
  q5 = new WeakSet();
function G5(e, t) {
  if (e.workbookLinks === undefined) return;
  if (
    e.workbookLinks &&
    typeof e.workbookLinks == "object" &&
    q5.has(e.workbookLinks)
  ) {
    t.workbookLinks = e.workbookLinks;
    return;
  }
  if (JSON.stringify(e.workbookLinks).length > V2.metadataCharacters)
    throw Error("Workbook link metadata exceeds 2 million characters.");
  let n = eW.parse(e.workbookLinks);
  if (new Set(n.sources.map((o) => o.target)).size !== n.sources.length)
    throw Error("Workbook link targets must be unique.");
  if (
    new Set(n.packageTargets?.map((o) => o.key)).size !==
    (n.packageTargets?.length ?? 0)
  )
    throw Error("Package link identities must be unique.");
  let i = 0,
    r = 0;
  for (let o of n.sources) {
    if (
      new Set(o.sheets.map((a) => a.name.toLowerCase())).size !==
      o.sheets.length
    )
      throw Error("Linked worksheet names must be unique.");
    for (let a of o.sheets) {
      r++;
      for (let s of Object.keys(a.cells)) (Zm(s), i++);
    }
  }
  if (i > V2.cells || r > V2.sheets)
    throw Error("Workbook link cache exceeds its cell or sheet limit.");
  ((t.workbookLinks = rl(n)), q5.add(t.workbookLinks));
}
function Z5(e) {
  let t = [],
    n = new Map(),
    i = new Map(),
    r = new Set(e.sheets.map((a) => a.name.toLowerCase()));
  for (let [a, s] of (e.workbookLinks?.sources ?? []).entries())
    for (let [l, c] of s.sheets.entries()) {
      let u = `_LG_LINK_${a}_${l}`;
      for (; r.has(u.toLowerCase());) u += "_";
      r.add(u.toLowerCase());
      let d = `linked:${a}:${l}`;
      for (; e.sheets.some((p) => p.id === d);) d += "_";
      let f = {};
      for (let [p, h] of Object.entries(c.cells))
        ((f[p] = { raw: h.value, literal: true }), n.set(`${d}:${p}`, h));
      (t.push({ id: d, name: u, rowCount: 1048576, colCount: 16384, cells: f }),
        i.set(JSON.stringify([s.target, c.name.toLowerCase()]), u));
    }
  return {
    sheets: t,
    values: n,
    rewrite: (a) =>
      V5(a, (s) => {
        let l = e.workbookLinks?.sources.find((u) => u.target === s.target),
          c = l?.sheets.find(
            (u) => u.name.toLowerCase() === s.sheet.toLowerCase(),
          );
        if (!(!c || !Zm(s.range).every((u) => Object.hasOwn(c.cells, u))))
          return `'${i.get(JSON.stringify([s.target, s.sheet.toLowerCase()]))}'!${s.range}`;
      }),
  };
}
function K5(e, t, n, i, r = true) {
  let o = "";
  for (let u = 0; u < e.length;) {
    let d = u,
      f = e[u];
    if (f === '"' || f === "'") {
      for (u++; u < e.length;)
        if (e[u++] === f)
          if (e[u] === f) u++;
          else break;
      o += f === '"' ? '""' : e.slice(d, u);
      continue;
    }
    let p = /^(?:[A-Za-z_][A-Za-z0-9_.]*)?\[/.exec(e.slice(u));
    if (p) {
      u += p[0].length;
      let h = 1;
      for (; u < e.length && h;)
        (e[u] === "[" && h++, e[u] === "]" && h--, u++);
      let m = e.slice(d, u);
      o += !h && Hm(t, n, m, i) ? "TABLETOKEN" : m;
      continue;
    }
    o += e[u++];
  }
  let a = (u, d, f) => ({
    sheetId: n,
    address: i,
    kind: u,
    error: d,
    message: f,
  });
  if (
    /\[[^\]]+\][^!+*/^&=<>(),;{}]*!/.test(o) ||
    /\|/.test(o) ||
    /(?:https?:|file:|\\\\)/i.test(o)
  )
    return a(
      "external",
      "#BLOCKED!",
      "External workbook, path, network and DDE references are not supported. Use an authorized local worksheet reference.",
    );
  let s = `(?:'(?:[^']|'')*'|[^\\s!+*/^&=<>(),;:{}\\[\\]\\"'%\\-]+)`;
  if (
    /'(?:[^']|'')*:(?:[^']|'')*'\s*!/.test(o) ||
    new RegExp(`${s}\\s*:\\s*${s}\\s*!`).test(o)
  )
    return a(
      "three-dimensional",
      "#VALUE!",
      "3-D references across a worksheet interval are not supported. List the individual worksheet references explicitly.",
    );
  let c = o
    .replace(/'(?:[^']|'')*'\s*!/g, "Sheet!")
    .replace(
      /(?<![\p{L}\p{N}_.\]])\$?[A-Za-z]{1,3}\$?[1-9]\d*(?![\p{L}\p{N}_.])/gu,
      "CELLREF",
    );
  if (
    r &&
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
var q2 = [
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
  X2 = Symbol("not a database function"),
  zs = {
    databaseCells: 2e4,
    criteriaCells: 1e3,
    predicateComparisons: 1e5,
    textComparisonSteps: 5e6,
    criterionCharacters: 255,
    headerCharacters: 255,
    formulaEvaluations: 1e4,
  };
function J5(e, t, n) {
  if (!q2.includes(e)) return X2;
  let i = (x = "#VALUE!") => {
    throw n.issue(x);
  };
  if (t.length !== 3 || t[0] == null || t[2] == null) return i();
  let r = (x, C) => {
      let N = n.retrieve(x);
      if (n.isIssue(N)) throw N;
      if (!Array.isArray(N) || N.length < 2) return i();
      if (N.length > C) return i("#LIMIT!");
      if (!Array.isArray(N[0])) return i();
      let F = N,
        z = F[0].length;
      if (!z) return i();
      if (F.length * z > C) return i("#LIMIT!");
      for (let j = 0; j < F.length; j++)
        if (!Array.isArray(F[j]) || F[j].length !== z) return i();
      for (let j of F)
        for (let q of j)
          if (n.isIssue(q)) {
            if (
              [
                "#LIMIT!",
                "#CYCLE!",
                "#BLOCKED!",
                "#TIMEOUT!",
                "#ITER!",
                "#ERROR!",
              ].includes(q.error)
            )
              throw q;
          } else {
            if (
              q != null &&
              !["number", "string", "boolean"].includes(typeof q)
            )
              return i();
            if (typeof q == "number" && !Number.isFinite(q)) return i("#NUM!");
          }
      return F;
    },
    o = r(t[0], zs.databaseCells),
    a = r(t[2], zs.criteriaCells);
  if ((o.length - 1) * (a.length - 1) * a[0].length > zs.predicateComparisons)
    return i("#LIMIT!");
  let s = (x) => {
      if (n.isIssue(x)) throw x;
      return typeof x != "string" || !x.trim() || x.length > zs.headerCharacters
        ? i()
        : x.toLowerCase();
    },
    l = o[0].map(s);
  if (new Set(l).size !== l.length) return i();
  let c = a[0].map((x) => {
      if (n.isIssue(x)) throw x;
      return x == null || x === "" ? -1 : l.indexOf(s(x));
    }),
    u = e === "DCOUNT" || e === "DCOUNTA",
    d;
  if (t[1] == null) {
    if (!u) return i();
  } else {
    let x = n.retrieve(t[1]);
    if (n.isIssue(x)) throw x;
    if (typeof x == "number") {
      if (!Number.isInteger(x) || x < 1 || x > l.length) return i();
      d = x - 1;
    } else if (typeof x == "string" && x.length) {
      if (((d = l.indexOf(x.toLowerCase())), d < 0)) return i();
    } else return i();
  }
  let f = 0,
    p = (x) => {
      ((f += x), f > zs.textComparisonSteps && i("#LIMIT!"));
    },
    h = (x, C, N) =>
      N === "<>"
        ? x !== C
        : N === "<"
          ? x < C
          : N === ">"
            ? x > C
            : N === "<="
              ? x <= C
              : N === ">="
                ? x >= C
                : x === C,
    m = (x) => {
      if (n.isIssue(x)) throw x;
      if (x == null || x === "") return null;
      if (typeof x == "number" || typeof x == "boolean") return (D) => D === x;
      if (typeof x != "string") return i();
      if (x.length > zs.criterionCharacters) return i("#LIMIT!");
      let C = /^(<=|>=|<>|=|<|>)?(.*)$/s.exec(x),
        N = C[1] ?? "=",
        F = C[2];
      if (C[1] && /^[<>=]/.test(F)) return i();
      if (!F && C[1])
        return N !== "=" && N !== "<>"
          ? i()
          : (D) => (N === "=" ? D == null || D === "" : D != null && D !== "");
      if (/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(F.trim())) {
        let D = Number(F);
        return Number.isFinite(D)
          ? (W) => (typeof W == "number" ? h(W, D, N) : N === "<>")
          : i("#NUM!");
      }
      if (/^(TRUE|FALSE)$/i.test(F)) {
        if (N !== "=" && N !== "<>") return i();
        let D = F.toUpperCase() === "TRUE";
        return (W) => (N === "=" ? W === D : W !== D);
      }
      let z = F.toLowerCase(),
        j = [],
        q = false;
      for (let D = 0; D < z.length; D++) {
        let W = z[D];
        if (W === "~") {
          let $ = z[D + 1];
          $ && ["*", "?", "~"].includes($)
            ? (j.push({ kind: "literal", char: $ }), D++, (q = true))
            : j.push({ kind: "literal", char: W });
        } else
          W === "*" || W === "?"
            ? ((q = true),
              (W === "?" || j.at(-1)?.kind !== "many") &&
                j.push({ kind: W === "*" ? "many" : "one" }))
            : j.push({ kind: "literal", char: W });
      }
      return ["<", ">", "<=", ">="].includes(N)
        ? q
          ? i()
          : (D) =>
              typeof D != "string"
                ? false
                : (p(D.length + z.length), h(D.toLowerCase(), z, N))
        : (!C[1] && !q && j.push({ kind: "many" }),
          (D) => {
            if (typeof D != "string") return N === "<>";
            p((D.length + 1) * (j.length + 1));
            let W = D.toLowerCase(),
              $ = new Uint8Array(W.length + 1);
            $[0] = 1;
            for (let B of j) {
              let V = new Uint8Array(W.length + 1);
              if (B.kind === "many") {
                V[0] = $[0];
                for (let G = 1; G <= W.length; G++) V[G] = $[G] || V[G - 1];
              } else
                for (let G = 1; G <= W.length; G++)
                  V[G] =
                    $[G - 1] && (B.kind === "one" || B.char === W[G - 1])
                      ? 1
                      : 0;
              $ = V;
            }
            let M = !!$[W.length];
            return N === "<>" ? !M : M;
          });
    },
    g = 0,
    b = a
      .slice(1)
      .map((x, C) =>
        x.map((N, F) =>
          c[F] < 0 ? n.criterionFormula?.(t[2], C + 1, F) : undefined,
        ),
      );
  c.forEach((x, C) => {
    x < 0 && !b.some((N) => N[C]) && i();
  });
  let w = a.slice(1).map((x, C) =>
    x.map((N, F) => {
      if (c[F] >= 0) return m(N);
      if (!b[C][F] && !(N == null || N === "")) return i();
    }),
  );
  b.flat().filter(Boolean).length * (o.length - 1) > zs.formulaEvaluations &&
    i("#LIMIT!");
  let S = 0,
    O = 0,
    E = 0,
    R = 0,
    I = 1,
    U = 1 / 0,
    P = -1 / 0,
    T = 0,
    A = 0,
    L;
  for (let [x, C] of o.slice(1).entries()) {
    let N = false;
    for (let [j, q] of w.entries()) {
      let D = true;
      for (let W = 0; W < q.length; W++) {
        let $ = b[j][W];
        if ($) {
          ++g > zs.formulaEvaluations && i("#LIMIT!");
          let V = $(x);
          if (n.isIssue(V)) throw V;
          (typeof V != "boolean" && i(), V || (D = false));
          continue;
        }
        let M = q[W];
        if (!M) continue;
        let B = C[c[W]];
        if (n.isIssue(B)) throw B;
        M(B) || (D = false);
      }
      D && (N = true);
    }
    if (!N || (S++, d === undefined)) continue;
    let F = C[d];
    if (((L = F), F != null && E++, n.isIssue(F))) {
      if (!["DCOUNT", "DCOUNTA", "DGET"].includes(e)) throw F;
      continue;
    }
    if (typeof F != "number") continue;
    (O++, (R += F), (I *= F), (U = Math.min(U, F)), (P = Math.max(P, F)));
    let z = F - T;
    ((T += z / O), (A += z * (F - T)));
  }
  if (e === "DGET") {
    if (S !== 1) return i(S ? "#NUM!" : "#VALUE!");
    if (n.isIssue(L)) throw L;
    return L ?? 0;
  }
  if (e === "DCOUNT") return d === undefined ? S : O;
  if (e === "DCOUNTA") return d === undefined ? S : E;
  let y = (x) => (Number.isFinite(x) ? x : i("#NUM!"));
  if (e === "DSUM") return y(R);
  if (e === "DMAX") return O ? P : 0;
  if (e === "DMIN") return O ? U : 0;
  if (e === "DPRODUCT") return O ? y(I) : 0;
  if (e === "DAVERAGE") return O ? y(R / O) : i("#DIV/0!");
  let v = e === "DSTDEV" || e === "DVAR";
  if (O < (v ? 2 : 1)) return i("#DIV/0!");
  let k = Math.max(0, A / (O - (v ? 1 : 0)));
  return y(e === "DSTDEV" || e === "DSTDEVP" ? Math.sqrt(k) : k);
}
var H2 = ["SUBTOTAL", "AGGREGATE", "ISFORMULA", "FORMULATEXT"],
  G2 = Symbol("not-reference-function"),
  Bs = (e) =>
    !!e &&
    typeof e == "object" &&
    !Array.isArray(e) &&
    [Object.prototype, null].includes(Object.getPrototypeOf(e)),
  Km = (e) => typeof e?.raw == "string" && e.raw.startsWith("=") && !e.literal;
function tW(e) {
  return /(?:^|[^A-Za-z0-9_.])(?:_xlfn\.(?:_xlws\.)?)?SUBTOTAL\s*\(/i.test(
    e.replace(/"(?:[^"]|"")*"|'(?:[^']|'')*'(?=\s*!)/g, ""),
  );
}
function rW(e) {
  return /(?:^|[^A-Za-z0-9_.])(?:_xlfn\.(?:_xlws\.)?)?(?:SUBTOTAL|AGGREGATE)\s*\(/i.test(
    e.replace(/"(?:[^"]|"")*"|'(?:[^']|'')*'(?=\s*!)/g, ""),
  );
}
function Y5(e, t = {}) {
  if (!Bs(t) || Object.keys(t).some((l) => l !== "filteredRowsBySheet"))
    throw Error("Calculation options support only filteredRowsBySheet.");
  let n = t.filteredRowsBySheet;
  if (n !== undefined && !Bs(n))
    throw Error(
      "Filtered rows must be a worksheet-ID map of inclusive row intervals.",
    );
  let i = new Map(e.sheets.map((l) => [l.id, l])),
    r = new Map(e.sheets.map((l) => [l.name.toLowerCase(), l])),
    o = new Map(),
    a = 0;
  for (let [l, c] of Object.entries(n ?? {})) {
    let u = i.get(l);
    if (!u) throw Error("Filter context refers to an unknown worksheet ID.");
    if (!Array.isArray(c) || (a += c.length) > 1e4)
      throw Error(
        "Filter context supports at most 10,000 intervals across the workbook.",
      );
    o.set(l, Oa(u.rowCount, Vo(c, u.rowCount)));
  }
  let s = new Map();
  for (let l of e.sheets) s.set(l.id, Oa(l.rowCount, qo(l, "row")));
  return (l, c, u) => {
    if (!H2.includes(l)) return G2;
    let { issue, isIssue } = u,
      p = (D) => {
        if (isIssue(D)) throw D;
        if (!Bs(D) || !Bs(D.ref)) throw issue("#VALUE!");
        let W = D.ref;
        if (W.sheet !== undefined && typeof W.sheet != "string")
          throw issue("#REF!");
        let $ = r.get(
          String(W.sheet ?? u.currentSheetName ?? "").toLowerCase(),
        );
        if (!$) throw issue("#REF!");
        let M = Bs(W.from) ? W.from : W,
          B = Bs(W.to) ? W.to : W,
          V = [M.row, M.col, B.row, B.col];
        if (V.some((Q) => !Number.isInteger(Q))) throw issue("#REF!");
        let [G, H, K, te] = V;
        if (
          G < 1 ||
          H < 1 ||
          K < G ||
          te < H ||
          K > $.rowCount ||
          te > $.colCount
        )
          throw issue("#REF!");
        return { sheet: $, firstRow: G, lastRow: K, firstCol: H, lastCol: te };
      };
    if (l === "FORMULATEXT" || l === "ISFORMULA") {
      if (c.length !== 1) return issue("#VALUE!");
      let D = p(c[0]),
        W = (M, B) => {
          let V = `${Ue(B - 1)}${M}`;
          return (u.readMetadata?.(D.sheet.id, V), D.sheet.cells[V]);
        };
      if (l === "FORMULATEXT") {
        let M = W(D.firstRow, D.firstCol);
        return Km(M) && M.raw.length <= 8192 ? M.raw : issue("#N/A");
      }
      let $ = (D.lastRow - D.firstRow + 1) * (D.lastCol - D.firstCol + 1);
      return $ > 1e4
        ? issue("#LIMIT!")
        : $ === 1
          ? Km(W(D.firstRow, D.firstCol))
          : Array.from({ length: D.lastRow - D.firstRow + 1 }, (M, B) =>
              Array.from({ length: D.lastCol - D.firstCol + 1 }, (V, G) =>
                Km(W(D.firstRow + B, D.firstCol + G)),
              ),
            );
    }
    let h = l === "AGGREGATE";
    if (c.length < (h ? 3 : 2) || c.length > 255) return issue("#VALUE!");
    let m = (D) => {
        let W = u.retrieve(D);
        if (isIssue(W)) throw W;
        if (
          Array.isArray(W) ||
          !["number", "string", "boolean"].includes(typeof W) ||
          (typeof W == "string" && W.trim() === "") ||
          !Number.isFinite(Number(W))
        )
          throw issue("#VALUE!");
        return Number(W);
      },
      g = m(c[0]);
    if (h && !Number.isInteger(g)) return issue("#VALUE!");
    let b = Math.trunc(g),
      w = !h && b > 100 ? b - 100 : b;
    if (w < 1 || w > (h ? 19 : 11)) return issue("#VALUE!");
    let _ = h ? (c[1] == null ? 0 : m(c[1])) : 0;
    if (!Number.isInteger(_) || _ < 0 || _ > 7) return issue("#VALUE!");
    let S = h ? _ % 2 === 1 : b > 100,
      O = h && [2, 3, 6, 7].includes(_),
      E = !h || _ < 4,
      R = h && w >= 14;
    if (R && (c.length !== 4 || c[3] == null)) return issue("#VALUE!");
    let I = R ? m(c[3]) : 0,
      U = R ? [c[2]] : c.slice(h ? 2 : 1),
      T = R && !((D) => Bs(D) && Bs(D.ref))(U[0]),
      A = T ? [] : U.map(p),
      L = 0;
    for (let D of A) {
      let W = (D.lastRow - D.firstRow + 1) * (D.lastCol - D.firstCol + 1);
      if (((L += W), W > 2e4 || L > 1e5)) return issue("#LIMIT!");
    }
    for (let D of A)
      for (let W = D.firstRow; W <= D.lastRow; W++)
        for (let $ = D.firstCol; $ <= D.lastCol; $++)
          if (u.isActive(D.sheet.id, `${Ue($ - 1)}${W}`))
            return issue("#CYCLE!");
    let y = 0,
      v = 0,
      k = 1 / 0,
      x = -1 / 0,
      C = 0,
      N = 0,
      F = [],
      z = (D) => {
        if (isIssue(D)) {
          if (["#LIMIT!", "#CYCLE!"].includes(D.error)) return D;
          if (O) return;
          if (w !== 2 && w !== 3) return D;
        }
        if ((D != null && v++, typeof D != "number")) return;
        if (!Number.isFinite(D)) return issue("#NUM!");
        (y++, (k = Math.min(k, D)), (x = Math.max(x, D)));
        let W = D - C;
        ((C += W / y),
          (N += W * (D - C)),
          (w === 1 || w === 6 || w === 9 || w >= 12) && F.push(D));
      };
    if (T) {
      let D = u.retrieve(U[0]),
        W = Array.isArray(D) ? D : [[D]];
      if (!W.length || W.length > 1e4) return issue("#LIMIT!");
      let $ = W.every(Array.isArray) ? W : [W],
        M = $[0]?.length ?? 0;
      if (!M || $.some((B) => B.length !== M || B.some(Array.isArray)))
        return issue("#VALUE!");
      if ($.length * M > 1e4) return issue("#LIMIT!");
      for (let B of $)
        for (let V of B) {
          if (
            V != null &&
            !isIssue(V) &&
            !["number", "string", "boolean"].includes(typeof V)
          )
            return issue("#VALUE!");
          let G = z(V);
          if (G) return G;
        }
    }
    for (let D of A)
      for (let W = D.firstRow; W <= D.lastRow; W++)
        if (!(
          o.get(D.sheet.id)?.isHidden(W - 1) ||
          (S && s.get(D.sheet.id).isHidden(W - 1))
        ))
          for (let $ = D.firstCol; $ <= D.lastCol; $++) {
            let M = `${Ue($ - 1)}${W}`;
            u.readMetadata?.(D.sheet.id, M);
            let B = D.sheet.cells[M];
            if (E && Km(B) && (h ? rW : tW)(u.expandedFormula(B, D.sheet.id)))
              continue;
            let V = z(u.readCell(D.sheet.id, M));
            if (V) return V;
          }
    if (w >= 12) {
      if ((F.sort((G, H) => G - H), w === 13)) {
        let G = 1,
          H = 0,
          K = 0,
          te;
        for (let Q of F)
          ((K = Q === te ? K + 1 : 1), K > G && ((G = K), (H = Q)), (te = Q));
        return G > 1 ? H : issue("#N/A");
      }
      if (!F.length) return issue("#NUM!");
      if (w === 12) {
        let G = Math.floor(F.length / 2);
        return F.length % 2 ? F[G] : F[G - 1] / 2 + F[G] / 2;
      }
      if (w === 14 || w === 15) {
        let G = I;
        return Number.isInteger(G) && G >= 1 && G <= F.length
          ? F[w === 14 ? F.length - G : G - 1]
          : issue("#NUM!");
      }
      let D = w >= 18,
        W = w === 17 || w === 19 ? Math.trunc(I) / 4 : I;
      if (W < 0 || W > 1 || (D && (W === 0 || W === 1))) return issue("#NUM!");
      let $ = D ? W * (F.length + 1) - 1 : W * (F.length - 1);
      if ($ < 0 || $ > F.length - 1) return issue("#NUM!");
      let M = Math.floor($),
        B = Math.ceil($),
        V = $ - M;
      return F[M] * (1 - V) + F[B] * V;
    }
    if (w === 2) return y;
    if (w === 3) return v;
    if (w === 9) return oi(F);
    if (w === 1) return y ? sn(F) : issue("#DIV/0!");
    if (w === 4) return y ? x : 0;
    if (w === 5) return y ? k : 0;
    if (w === 6) return y ? Nm(F) : 0;
    let j = w === 7 || w === 10;
    if (y < (j ? 2 : 1)) return issue("#DIV/0!");
    let q = Math.max(0, N / (y - (j ? 1 : 0)));
    return w === 7 || w === 8 ? Math.sqrt(q) : q;
  };
}
var nW = [
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
  Pm = [
    ...new Set([
      ...nW,
      ...bp,
      ...k2,
      ...Cw,
      ...H2,
      ...q2,
      ...y2,
      ...GI,
      ...l5,
      ...m2,
      ...d2,
      ...h2,
    ]),
  ].sort(),
  iW = new Set(Pm),
  Q5 = K2.default.FormulaError,
  J2 = (e) => new Q5(e),
  oW = J2,
  dt = (e) => e instanceof Q5,
  Cn = (e) => (Array.isArray(e) ? e.flatMap(Cn) : [e]),
  Z2 = (e) => e.flatMap(Cn).find(dt);
function Od(e, t) {
  if (typeof t != "string") return e === t;
  let n = /^(<=|>=|<>|=|<|>)?(.*)$/.exec(t),
    i = n[1] ?? "=",
    r = n[2];
  if ((i === "=" || i === "<>") && /[?*~]/.test(r)) {
    if (typeof e != "string") return i === "<>";
    let l = [];
    for (let d = 0; d < r.length; d++)
      r[d] === "~" && d + 1 < r.length
        ? l.push({ char: r[++d].toLowerCase(), wildcard: false })
        : l.push({
            char: r[d].toLowerCase(),
            wildcard: r[d] === "*" || r[d] === "?",
          });
    let c = String(e ?? "").toLowerCase();
    if (l.length * c.length > 1e5) throw oW("#LIMIT!");
    let u = Array(c.length + 1).fill(false);
    u[0] = true;
    for (let d of l) {
      let f = Array(c.length + 1).fill(false);
      f[0] = d.wildcard && d.char === "*" && u[0];
      for (let p = 1; p <= c.length; p++)
        f[p] =
          d.wildcard && d.char === "*"
            ? u[p] || f[p - 1]
            : u[p - 1] &&
              ((d.wildcard && d.char === "?") || c[p - 1] === d.char);
      u = f;
    }
    return i === "<>" ? !u[c.length] : u[c.length];
  }
  let o = r.trim() !== "" && Number.isFinite(Number(r)),
    a = o
      ? e === null || e === ""
        ? 0
        : typeof e == "number"
          ? e
          : NaN
      : String(e ?? "").toLowerCase(),
    s = o ? Number(r) : r.toLowerCase();
  return i === "<>"
    ? a !== s
    : i === "<"
      ? a < s
      : i === ">"
        ? a > s
        : i === "<="
          ? a <= s
          : i === ">="
            ? a >= s
            : a === s;
}
function sW(e, t, n = J2, i = zu(), r, o) {
  let a = e.utils._applyInfix.bind(e.utils);
  ((e.utils._applyInfix = (l, c, u) => {
    let d = (b, w) => (
      ["=", "<>", "<", ">", "<=", ">="].includes(c) &&
        typeof b == "string" &&
        typeof w == "string" &&
        ((b = b.toLowerCase()), (w = w.toLowerCase())),
      a({ val: b, isArray: false }, c, { val: w, isArray: false })
    );
    if (!Array.isArray(l.val) && !Array.isArray(u.val)) return d(l.val, u.val);
    let f = (b) =>
        Array.isArray(b) ? (b.every(Array.isArray) ? b : [b]) : [[b]],
      p = f(l.val),
      h = f(u.val),
      m = Math.max(p.length, h.length),
      g = Math.max(p[0]?.length ?? 0, h[0]?.length ?? 0);
    return !m ||
      !g ||
      (p.length !== m && p.length !== 1) ||
      (h.length !== m && h.length !== 1) ||
      (p[0].length !== g && p[0].length !== 1) ||
      (h[0].length !== g && h[0].length !== 1)
      ? n("#VALUE!")
      : m * g > dr.maxSpillCells
        ? n("#LIMIT!")
        : Array.from({ length: m }, (b, w) =>
            Array.from({ length: g }, (_, S) =>
              d(
                p[p.length === 1 ? 0 : w][p[0].length === 1 ? 0 : S],
                h[h.length === 1 ? 0 : w][h[0].length === 1 ? 0 : S],
              ),
            ),
          );
  }),
    (e.checkFormulaResult = (l, c) => {
      let u = e.retrieveRef(l);
      return dt(u)
        ? u
        : Array.isArray(u)
          ? c
            ? u
            : n("#VALUE!")
          : u === null
            ? 0
            : typeof u == "number"
              ? Number.isFinite(u)
                ? u
                : n("#NUM!")
              : typeof u == "string" || typeof u == "boolean"
                ? u
                : n("#VALUE!");
    }));
  let s = e._callFunction.bind(e);
  return (
    (e._callFunction = (l, c) => {
      l = l
        .replace(/^_xlfn\./i, "")
        .replace(/^_xlws\./i, "")
        .toUpperCase();
      let u = (d) => (d == null ? undefined : e.retrieveRef(d));
      try {
        if (!iW.has(l)) return n("#NAME?");
        if (l === "RAND") return c.length ? n("#VALUE!") : Math.random();
        if (l === "RANDBETWEEN") {
          if (c.length !== 2) return n("#VALUE!");
          let A = c.map(u),
            L = Z2(A);
          if (L) return L;
          if (
            A.some(
              (k) =>
                Array.isArray(k) ||
                k === undefined ||
                (typeof k == "string" && k.trim() === "") ||
                !Number.isFinite(Number(k)),
            )
          )
            return n("#VALUE!");
          let y = Math.ceil(Number(A[0])),
            v = Math.floor(Number(A[1]));
          return !Number.isSafeInteger(y) ||
            !Number.isSafeInteger(v) ||
            y > v ||
            !Number.isSafeInteger(v - y + 1)
            ? n("#NUM!")
            : y + Math.floor(Math.random() * (v - y + 1));
        }
        let d = Bu(l, c, { retrieve: u, issue: n, isIssue: dt, clock: i });
        if (d !== Ew) return d;
        let f = P6(l, c, {
          retrieve: u,
          issue: n,
          isIssue: dt,
          clock: i,
          report: r,
        });
        if (f !== Tw) return f;
        let p = t(l, c, e);
        if (p !== G2) return p;
        let h = J5(l, c, {
          retrieve: u,
          issue: n,
          isIssue: dt,
          criterionFormula: o,
        });
        if (h !== X2) return h;
        let m = XI(l, c, { retrieve: u, issue: n, isIssue: dt });
        if (m !== g2) return m;
        let g = jI(l, c, { retrieve: u, issue: n, isIssue: dt });
        if (g !== f2) return g;
        let b = WI(l, c, {
          retrieve: u,
          issue: n,
          isIssue: dt,
          distribution: s,
        });
        if (b !== p2) return b;
        let w = BI(l, c, { retrieve: u, issue: n, isIssue: dt });
        if (w !== u2) return w;
        let _ = HI(l, c, { retrieve: u, issue: n, isIssue: dt });
        if (_ !== v2) return _;
        let S = KI(l, c, { retrieve: u, issue: n, isIssue: dt });
        if (S !== x2) return S;
        let O = ZI(l, c, { retrieve: u, issue: n, isIssue: dt, matches: Od });
        if (O !== w2) return O;
        if (l === "TEXTBEFORE" || l === "TEXTAFTER")
          return Sp(
            c.map(u),
            (A) => {
              try {
                return Iw(l, A, {
                  retrieve: (L) => L,
                  issue: n,
                  isIssue: dt,
                  format: () => n("#VALUE!"),
                });
              } catch (L) {
                return dt(L) ? L : n("#VALUE!");
              }
            },
            n,
          );
        let E = Iw(l, c, {
          retrieve: u,
          issue: n,
          isIssue: dt,
          temporal: (A) => Hi(A, i),
          format: () => {
            let A = Number(u(c[0]) ?? 0),
              L = u(c[1]);
            if (typeof A == "number" && typeof L == "string") {
              try {
                if (Ju(L).kind === "temporal")
                  return oC(
                    A,
                    { category: "custom", locale: "en-US", pattern: L },
                    i.dateSystem,
                  ).text;
              } catch {}
              let y = L.replace(/"[^"]*"|\\./g, "");
              if (
                i.dateSystem === "1904" &&
                /[ydm]/i.test(y) &&
                !/\[[hms]+\]/i.test(y)
              )
                return s(l, [A + ka(i.dateSystem), L]);
            }
            return s(l, c);
          },
        });
        if (E !== Cp) return E;
        let R = QI(l, c, { retrieve: u, issue: n, isIssue: dt, matches: Od });
        if (R !== $m) return R;
        if (l === "INDEX") {
          if (c.length < 2 || c.length > 4) return n("#VALUE!");
          if (dt(c[0])) return c[0];
          let A = [...c];
          for (let y = 1; y < c.length; y++) {
            let v = c[y] == null ? (y === 3 ? 1 : 0) : u(c[y]);
            if (dt(v)) return v;
            if (
              Array.isArray(v) ||
              (typeof v == "string" && v.trim() === "") ||
              !Number.isFinite(Number(v)) ||
              Number(v) < 0
            )
              return n("#VALUE!");
            A[y] = Number(v);
          }
          if (A.length > 3 && Number(A[3]) < 1) return n("#VALUE!");
          let L = s(l, A);
          if (L && typeof L == "object" && "ref" in L) {
            let y = c[0],
              v = L;
            y?.ref?.sheet && !v.ref.sheet && (v.ref.sheet = y.ref.sheet);
          }
          return L;
        }
        if (l === "IF") {
          if (c.length < 2 || c.length > 3) return n("#VALUE!");
          let A = u(c[0]);
          if (dt(A)) return A;
          if (Array.isArray(A))
            return Sp(
              [A, u(c[1]), c.length === 3 ? u(c[2]) : false],
              ([y, v, k]) =>
                dt(y)
                  ? y
                  : typeof y == "string" && !/^(TRUE|FALSE)$/i.test(y)
                    ? n("#VALUE!")
                    : (((
                        typeof y == "string" ? y.toUpperCase() === "TRUE" : !!y
                      )
                        ? v
                        : k) ?? 0),
              n,
            );
          if (typeof A == "string" && !/^(TRUE|FALSE)$/i.test(A))
            return n("#VALUE!");
          let L = typeof A == "string" ? A.toUpperCase() === "TRUE" : !!A;
          return u(L ? c[1] : c.length === 3 ? c[2] : false) ?? 0;
        }
        if (l === "IFERROR" || l === "IFNA") {
          if (c.length !== 2) return n("#VALUE!");
          let A = u(c[0]);
          return Array.isArray(A)
            ? Sp(
                [A, u(c[1])],
                ([L, y]) =>
                  dt(L) && (l === "IFERROR" || L.error === "#N/A")
                    ? (y ?? 0)
                    : (L ?? 0),
                n,
              )
            : dt(A) && (l === "IFERROR" || A.error === "#N/A")
              ? (u(c[1]) ?? 0)
              : (A ?? 0);
        }
        let I = c.map(u),
          U = I.flatMap(Cn).filter((A) => typeof A == "number");
        if (
          I.flatMap(Cn).length > 1e5 ||
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
          ].includes(l) &&
            ((Math.abs(Number(I[l === "BASE" ? 2 : 1] ?? 0)) > 32767 &&
              !["FIXED", "DOLLAR"].includes(l)) ||
              (["FIXED", "DOLLAR"].includes(l) &&
                Math.abs(Number(I[1] ?? 2)) > 100) ||
              (l === "REPT" &&
                String(I[0] ?? "").length * Number(I[1]) > 32767))) ||
          (["MUNIT", "MDETERM", "MMULT", "GROWTH"].includes(l) &&
            (l === "MUNIT"
              ? Number(I[0]) > 100
              : I.flatMap(Cn).length > 1e3)) ||
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
          ].includes(l) &&
            U.some((A) => Math.abs(A) > 1e4)) ||
          (["BESSELI", "BESSELJ", "BESSELK", "BESSELY"].includes(l) &&
            Math.abs(Number(I[1])) > 1e3)
        )
          return n("#LIMIT!");
        if (l === "ROMAN" && Math.abs(Number(I[0])) > 3999) return n("#VALUE!");
        if (
          [
            "NETWORKDAYS",
            "NETWORKDAYS.INTL",
            "WORKDAY",
            "WORKDAY.INTL",
          ].includes(l) &&
          (Math.abs(
            Number(I[1]) - (l.startsWith("NETWORK") ? Number(I[0]) : 0),
          ) > 366e3 ||
            I.flatMap(Cn).length > 1e3)
        )
          return n("#LIMIT!");
        if (l === "ISERR") return dt(I[0]) && I[0].error !== "#N/A";
        if (l === "ISLOGICAL") return typeof I[0] == "boolean";
        if (l === "ISNONTEXT") return typeof I[0] != "string";
        if (l === "N")
          return dt(I[0]) || typeof I[0] == "number"
            ? I[0]
            : I[0] === true
              ? 1
              : 0;
        if (l === "T") return dt(I[0]) || typeof I[0] == "string" ? I[0] : "";
        if (l === "TYPE")
          return dt(I[0])
            ? 16
            : Array.isArray(I[0])
              ? 64
              : typeof I[0] == "string"
                ? 2
                : typeof I[0] == "boolean"
                  ? 4
                  : 1;
        if (l === "ERROR.TYPE") {
          let A = [
            "#NULL!",
            "#DIV/0!",
            "#VALUE!",
            "#REF!",
            "#NAME?",
            "#NUM!",
            "#N/A",
            "#GETTING_DATA",
          ].indexOf(dt(I[0]) ? I[0].error : "");
          return A < 0 ? n("#N/A") : A + 1;
        }
        if (l === "ISERROR") return dt(I[0]);
        if (l === "ISNA") return dt(I[0]) && I[0].error === "#N/A";
        if (l === "ISNUMBER") return typeof I[0] == "number";
        if (l === "ISTEXT") return typeof I[0] == "string";
        if (l === "ISBLANK") return I[0] === null;
        if (l === "COUNTA")
          return I.flatMap(Cn).filter((A) => A !== null).length;
        if (l === "COUNTBLANK")
          return I.flatMap(Cn).filter((A) => A === null || A === "").length;
        if (
          [
            "SUMIF",
            "AVERAGEIF",
            "COUNTIF",
            "SUMIFS",
            "AVERAGEIFS",
            "COUNTIFS",
          ].includes(l)
        ) {
          let A = l.endsWith("IFS"),
            L = l.startsWith("COUNT"),
            y = l.startsWith("AVERAGE"),
            v,
            k;
          if (A) {
            if (((k = L ? I : I.slice(1)), k.length < 2 || k.length % 2 !== 0))
              return n("#VALUE!");
            v = L ? Cn(k[0]).map(() => 1) : Cn(I[0]);
          } else {
            if (I.length < 2 || I.length > (L ? 2 : 3)) return n("#VALUE!");
            ((k = I.slice(0, 2)),
              (v = L
                ? Cn(I[0]).map(() => 1)
                : Cn(c.length < 3 || c[2] == null ? I[0] : I[2])));
          }
          let x = [],
            C = [],
            N = (q) => {
              if (!Array.isArray(q)) return "1:1";
              let D = q.every(Array.isArray) ? q : [q],
                W = D[0]?.length ?? 0;
              return D.length && W && D.every(($) => $.length === W)
                ? `${D.length}:${W}`
                : "invalid";
            },
            F = N(L ? k[0] : A || c.length < 3 || c[2] == null ? I[0] : I[2]);
          for (let q = 0; q < k.length; q += 2) {
            let D = Cn(k[q]);
            if (
              D.length !== v.length ||
              F === "invalid" ||
              N(k[q]) !== F ||
              Array.isArray(k[q + 1])
            )
              return n("#VALUE!");
            if (dt(k[q + 1])) return k[q + 1];
            x.push(D);
            let W = k[q + 1];
            C.push(W == null ? "=0" : typeof W == "number" ? `=${W}` : W);
          }
          let z = [],
            j = 0;
          for (let q = 0; q < v.length; q++) {
            if (
              ((j += x.reduce(
                (D, W, $) =>
                  D +
                  Math.max(1, String(W[q] ?? "").length) *
                    Math.max(1, String(C[$] ?? "").length),
                0,
              )),
              j > 2e6)
            )
              return n("#LIMIT!");
            if (x.every((D, W) => !dt(D[q]) && Od(D[q], C[W]))) {
              if (dt(v[q])) return v[q];
              typeof v[q] == "number" && z.push(v[q]);
            }
          }
          return y ? (z.length ? sn(z) : n("#DIV/0!")) : oi(z);
        }
        let P = Z2(I);
        if (P && l !== "COUNT") return P;
        if (l === "TRIM")
          return String(I[0] ?? "")
            .replace(/^ +| +$/g, "")
            .replace(/ +/g, " ");
        if (l === "UPPER") {
          let A = String(I[0] ?? "").toUpperCase();
          return A.length > 32767 ? n("#LIMIT!") : A;
        }
        if (l === "VALUE") {
          if (I.length !== 1 || Array.isArray(I[0])) return n("#VALUE!");
          let A = String(I[0] ?? "").trim(),
            L = false;
          A.startsWith("(") &&
            A.endsWith(")") &&
            ((L = true), (A = A.slice(1, -1).trim()));
          let y = A.endsWith("%");
          if (
            (y && (A = A.slice(0, -1).trim()),
            (A = A.replace(/^\$/, "")),
            !/^[+-]?(?:\d{1,3}(?:,\d{3})+|\d+|\d*\.\d+)(?:\.\d*)?(?:[eE][+-]?\d+)?$/.test(
              A,
            ))
          )
            return n("#VALUE!");
          let v = (Number(A.replace(/,/g, "")) * (L ? -1 : 1)) / (y ? 100 : 1);
          return Number.isFinite(v) ? v : n("#NUM!");
        }
        if (l === "MATCH") {
          if (I.length < 2 || I.length > 3) return n("#VALUE!");
          let A = I[1];
          if (
            Array.isArray(A) &&
            A.length > 1 &&
            Array.isArray(A[0]) &&
            A[0].length > 1
          )
            return n("#N/A");
          let L = Cn(A),
            y = I[0],
            v = I[2] === undefined ? 1 : Number(I[2]);
          if (![0, -1, 1].includes(v)) return n("#N/A");
          if (v === 0) {
            let F = L.findIndex((z) =>
              typeof y == "string"
                ? typeof z == "string" &&
                  (/[?*~]/.test(y)
                    ? Od(z, `=${y}`)
                    : z.toLowerCase() === y.toLowerCase())
                : z === y,
            );
            return F < 0 ? n("#N/A") : F + 1;
          }
          let k = L.map((F) => (typeof F == "string" ? F.toLowerCase() : F)),
            x = typeof y == "string" ? y.toLowerCase() : y;
          if (
            (typeof x != "number" && typeof x != "string") ||
            k.some((F) => typeof F != typeof x)
          )
            return n("#VALUE!");
          let C = k;
          if (
            C.some((F, z) => z > 0 && (v === 1 ? F < C[z - 1] : F > C[z - 1]))
          )
            return n("#VALUE!");
          let N = -1;
          return (
            C.forEach((F, z) => {
              (v === 1 ? F <= x : F >= x) && (N = z);
            }),
            N < 0 ? n("#N/A") : N + 1
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
          ].includes(l)
        ) {
          let A = l === "LARGE" || l === "SMALL" ? I.slice(0, 1) : I,
            L = [];
          A.forEach((k, x) => {
            let C = !!c[x] && typeof c[x] == "object" && "ref" in c[x];
            for (let N of Cn(k))
              typeof N == "number"
                ? L.push(N)
                : !C &&
                  !Array.isArray(k) &&
                  (typeof N == "boolean" ||
                    (typeof N == "string" &&
                      N.trim() !== "" &&
                      Number.isFinite(Number(N)))) &&
                  L.push(Number(N));
          });
          let y = oi(L);
          if (l === "SUM") return y;
          if (l === "AVERAGE") return L.length ? y / L.length : n("#DIV/0!");
          if (l === "MIN")
            return L.length ? L.reduce((k, x) => Math.min(k, x)) : 0;
          if (l === "MAX")
            return L.length ? L.reduce((k, x) => Math.max(k, x)) : 0;
          if (l === "STDEV.S" || l === "VAR.S") {
            if (L.length < 2) return n("#DIV/0!");
            let k =
              L.reduce((x, C) => x + (C - y / L.length) ** 2, 0) /
              (L.length - 1);
            return l === "VAR.S" ? k : Math.sqrt(k);
          }
          if ((L.sort((k, x) => k - x), l === "MEDIAN"))
            return L.length
              ? (L[Math.floor((L.length - 1) / 2)] +
                  L[Math.floor(L.length / 2)]) /
                  2
              : n("#NUM!");
          let v = Number(I[1]);
          return !Number.isInteger(v) || v < 1 || v > L.length
            ? n("#NUM!")
            : l === "LARGE"
              ? L[L.length - v]
              : L[v - 1];
        }
        if (l === "SUBSTITUTE") {
          if (I.length < 3 || I.length > 4 || I.some(Array.isArray))
            return n("#VALUE!");
          let A = String(I[0] ?? ""),
            L = String(I[1] ?? ""),
            y = String(I[2] ?? ""),
            v = L.length ? A.split(L).length - 1 : 0;
          if (
            I[3] !== undefined &&
            (!Number.isInteger(Number(I[3])) || Number(I[3]) < 1)
          )
            return n("#VALUE!");
          if (
            A.length +
              (I[3] === undefined ? v : Math.min(v, 1)) *
                (y.length - L.length) >
            32767
          )
            return n("#LIMIT!");
          if (!L.length) return A;
          if (I[3] === undefined) return A.split(L).join(y);
          let k = 0;
          return A.split(L).reduce(
            (x, C, N) => (N === 0 ? C : x + (++k === Number(I[3]) ? y : L) + C),
            "",
          );
        }
        if (l === "SEARCH" || l === "SEARCHB") {
          if (I.length < 2 || I.length > 3 || I.some(Array.isArray))
            return n("#VALUE!");
          let A = String(I[0] ?? "").toLowerCase(),
            L = String(I[1] ?? "").toLowerCase(),
            y = I[2] === undefined ? 1 : Number(I[2]);
          if (!Number.isInteger(y) || y < 1 || y > L.length)
            return n("#VALUE!");
          if (!/[?*~]/.test(A)) {
            let v = L.indexOf(A, y - 1);
            return v < 0 ? n("#VALUE!") : v + 1;
          }
          if (L.length * L.length * A.length > 2e6) return n("#LIMIT!");
          for (let v = y - 1; v < L.length; v++)
            if (Od(L.slice(v), `${A}*`)) return v + 1;
          return n("#VALUE!");
        }
        if (
          l === "CONCATENATE" &&
          I.flatMap(Cn).reduce((A, L) => A + String(L ?? "").length, 0) > 32767
        )
          return n("#LIMIT!");
        let T = s(l, c);
        return /^IM/.test(l) &&
          !["IMABS", "IMREAL", "IMAGINARY", "IMARGUMENT"].includes(l) &&
          typeof T == "number"
          ? String(T)
          : typeof T == "string" && T.length > 32767
            ? n("#LIMIT!")
            : T;
      } catch (d) {
        return dt(d) ? n(d.error) : n("#ERROR!");
      }
    }),
    e
  );
}
function Hu(e, t = {}, n, i) {
  let r = kw(e),
    o = new Set([...Pm, ...r.map((V) => V.name)]),
    a = i?.budget ?? xp(e),
    s = i?.dateContext ?? zu(e.datePolicy),
    l = new Set(),
    c = new Map(),
    u = new Map(),
    d = new Map(),
    f = (V) => (
      [
        "#LIMIT!",
        "#BLOCKED!",
        "#CYCLE!",
        "#ERROR!",
        "#TIMEOUT!",
        "#ITER!",
      ].includes(V) && l.add(V),
      J2(V)
    ),
    p = Y5(e, t),
    h = new Map(),
    m = new Map(),
    g = new Map(),
    b = new Map(),
    w = new Set(),
    _ = new Set(),
    S = new Map(),
    O = [],
    E = [],
    R = 0,
    I = [],
    U = 0,
    P = e.definedNames?.some((V) =>
      /\b(?:_xlfn\.)?(?:LET|LAMBDA)\s*\(/i.test(V.refersTo),
    ),
    T = e.definedNames?.length
      ? new RegExp(
          `(?:^|[^A-Za-z0-9_.\\\\])(?:${e.definedNames.map((V) => V.name.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")).join("|")})(?![A-Za-z0-9_.\\\\])`,
          "i",
        )
      : undefined,
    A = (V) =>
      /\b(?:_xlfn\.)?(?:LET|LAMBDA)\s*\(/i.test(V) || !!(P && T?.test(V)),
    L = Z5(e),
    y = [...e.sheets, ...L.sheets],
    v = new Map(y.map((V) => [V.id, V])),
    k = new Map(y.map((V) => [V.name.toLowerCase(), V])),
    x = 0,
    C = 0,
    N = new Map(),
    F = new RegExp(
      `\\b(?:${[...kd, "ISFORMULA", "TEXTBEFORE", "TEXTAFTER", ...r.filter((V) => V.fn.result.type === "matrix").map((V) => V.name.replaceAll(".", "\\."))].join("|")})\\s*\\(`,
      "i",
    );
  for (let V of e.sheets)
    N.set(
      V.id,
      Object.entries(V.cells)
        .filter(
          ([, G]) =>
            typeof G.raw == "string" &&
            G.raw.startsWith("=") &&
            !G.literal &&
            (A(G.raw) ||
              F.test(G.raw) ||
              (/\b(?:IF|IFERROR|IFNA)\s*\(/i.test(G.raw) &&
                G.raw.replace(/"(?:[^"]|"")*"/g, '""').includes(":")) ||
              /[{}#]/.test(G.raw.replace(/"(?:[^"]|"")*"/g, '""')) ||
              /^=\s*[^()+*/]*:/.test(G.raw) ||
              (e.definedNames?.length &&
                /^=\s*[A-Za-z_][A-Za-z0-9_.]*\s*$/.test(G.raw))),
        )
        .map(([G]) => G)
        .sort((G, H) => {
          let K = tt(G),
            te = tt(H);
          return K.row - te.row || K.col - te.col;
        }),
    );
  let z = (V, G, H, K = true) => {
      let te = (ee, re = true) => {
        let oe = K5(ee, e, G, H, re);
        if (oe) throw (c.set(`${G}:${H}`, oe), f(oe.error));
      };
      ((V = L.rewrite(V)), te(V, false));
      let Q = K ? Ed(V, e, G, [], n?.allowName) : V;
      ((Q = L.rewrite(Q)), te(Q));
      let Z = "",
        J = 0;
      for (; J < Q.length;) {
        let ee =
          /^(?:'((?:[^']|'')*)'|([A-Za-z_][A-Za-z0-9_.]*))!\s*(\$?[A-Za-z]{1,3}\$?[1-9]\d*)#/i.exec(
            Q.slice(J),
          );
        if (ee) {
          let ne = k.get((ee[1]?.replace(/''/g, "'") ?? ee[2]).toLowerCase());
          if (!ne) throw f("#REF!");
          let ae = tt(ee[3]),
            le = `${Ue(ae.col)}${ae.row + 1}`,
            de = W(ne.id, le),
            Se = g.get(`${ne.id}:${le}`);
          if (!Se) throw dt(de) ? de : f("#REF!");
          ((Z += `'${ne.name.replace(/'/g, "''")}'!${Se}`),
            (J += ee[0].length));
          continue;
        }
        if (Q[J] === '"' || Q[J] === "'") {
          let ne = Q[J],
            ae = J++;
          for (; J < Q.length;)
            if (Q[J++] === ne)
              if (Q[J] === ne) J++;
              else break;
          Z += Q.slice(ae, J);
          continue;
        }
        let re =
          /^(SEQUENCE|FILTER|SORT|SORTBY|UNIQUE|TRANSPOSE|TAKE|DROP|TEXTSPLIT)\s*\(\s*,/i.exec(
            Q.slice(J),
          );
        if (re) {
          ((Z += `${re[1]}(${re[1].toUpperCase() === "SEQUENCE" ? "1" : "#VALUE!"},`),
            (J += re[0].length));
          continue;
        }
        let oe = /^(?:[A-Za-z_][A-Za-z0-9_.]*)?\[/.exec(Q.slice(J));
        if (oe) {
          let ne = J;
          J += oe[0].length;
          let ae = 1;
          for (; J < Q.length && ae;)
            (Q[J] === "[" && ae++, Q[J] === "]" && ae--, J++);
          let le = !ae && Hm(e, G, Q.slice(ne, J), H);
          if (!le) throw f("#REF!");
          let de = v.get(le.sheetId);
          Z += `'${de.name.replace(/'/g, "''")}'!${le.range}`;
          continue;
        }
        let Y = /^(\$?[A-Za-z]{1,3}\$?[1-9]\d*)#/.exec(Q.slice(J));
        if (Y) {
          let ne = tt(Y[1]),
            ae = `${Ue(ne.col)}${ne.row + 1}`,
            le = W(G, ae),
            de = g.get(`${G}:${ae}`);
          if (!de) throw dt(le) ? le : f("#REF!");
          ((Z += de), (J += Y[0].length));
          continue;
        }
        Z += Q[J++];
      }
      return Z.replace(/\b_xlfn\.(?:_xlws\.)?/gi, "");
    },
    j = (V) => {
      let G = k.get((V ?? "").toLowerCase());
      if (!G) throw f("#REF!");
      return G;
    },
    q = () => {
      let V = sW(
        new K2.default({
          onVariable: () => null,
          onCell: ({ sheet, row, col }) => {
            let re = j(sheet);
            if (
              !Number.isInteger(row) ||
              !Number.isInteger(col) ||
              row < 1 ||
              col < 1 ||
              row > re.rowCount ||
              col > re.colCount
            )
              throw f("#REF!");
            return W(re.id, `${Ue(col - 1)}${row}`);
          },
          onRange: ({ sheet, from, to }) => {
            let re = j(sheet),
              oe = to.row,
              Y = to.col;
            if (
              from.row < 1 ||
              from.col < 1 ||
              oe > re.rowCount ||
              Y > re.colCount ||
              oe < from.row ||
              Y < from.col
            )
              throw f("#REF!");
            if ((oe - from.row + 1) * (Y - from.col + 1) > 2e4)
              throw f("#LIMIT!");
            let ne = a.reserve(
              (oe - from.row + 1) * (32 + (Y - from.col + 1) * 8),
            );
            if (ne) throw f(ne);
            let ae = [];
            for (let le = from.row; le <= oe; le++) {
              let de = [];
              for (let Se = from.col; Se <= Y; Se++) {
                let ge = W(re.id, `${Ue(Se - 1)}${le}`);
                if (
                  dt(ge) &&
                  ["#LIMIT!", "#TIMEOUT!", "#BLOCKED!"].includes(ge.error)
                )
                  throw ge;
                de.push(ge);
              }
              ae.push(de);
            }
            return ae;
          },
        }),
        (Z, J, ee) =>
          p(Z, J, {
            currentSheetName: ee.position?.sheet,
            retrieve: (re) => ee.retrieveRef(re),
            issue: f,
            isIssue: dt,
            readCell: W,
            readMetadata:
              n || i
                ? (re, oe) => {
                    if ((i?.read(`${re}:${oe}`), n && !n.allowRead(re, oe)))
                      throw f("#BLOCKED!");
                    n?.step({
                      kind: "metadata",
                      sheetId: re,
                      address: oe,
                      operation: "Inspect cell metadata",
                      result: null,
                    });
                  }
                : undefined,
            isActive: (re, oe) => w.has(`${re}:${oe}`),
            expandedFormula: (re, oe) => {
              try {
                return Ed(String(re.raw), e, oe);
              } catch {
                return String(re.raw);
              }
            },
          }),
        f,
        s,
        (Z) => {
          let J = V.position;
          if (!J) return;
          let ee = j(J.sheet).id,
            re = `${Ue(J.col - 1)}${J.row}`,
            oe = `${ee}:${re}:${Z.functionName}`;
          if (d.size >= 1e4 && !d.has(oe)) throw f("#LIMIT!");
          d.set(oe, { ...Z, sheetId: ee, address: re });
        },
        (Z, J, ee) => {
          let re = Z && typeof Z == "object" && "ref" in Z ? Z.ref : undefined;
          if (!re?.from || !re.to) return;
          let oe = j(re.sheet ?? V.position?.sheet),
            Y = `${Ue(re.from.col + ee - 1)}${re.from.row + J}`;
          if (n && !n.allowRead(oe.id, Y)) throw f("#BLOCKED!");
          i?.read(`${oe.id}:${Y}`);
          let ne = oe.cells[Y];
          if (
            ne?.literal ||
            typeof ne?.raw != "string" ||
            !ne.raw.startsWith("=")
          )
            return;
          let ae = ne.raw;
          return (le) => {
            if (U >= 32) throw f("#LIMIT!");
            let de = U++;
            try {
              if (
                n?.allowDatabaseCriterion &&
                !n.allowDatabaseCriterion(oe.id, bi(ae, le, 0))
              )
                throw f("#BLOCKED!");
              let Se = Ed(ae, e, oe.id, [], n?.allowName),
                ge = bi(Se, le, 0);
              if (ge.length > dr.maxFormulaLength) throw f("#LIMIT!");
              return A(ge)
                ? D(ge.slice(1), oe.id, Y)
                : (I[de] ??= q()).parse(
                    z(ge.slice(1), oe.id, Y),
                    {
                      sheet: oe.name,
                      row: re.from.row + J,
                      col: re.from.col + ee,
                    },
                    true,
                  );
            } finally {
              U--;
            }
          };
        },
      );
      if (r.length) {
        let Z = V._callFunction.bind(V);
        V._callFunction = (J, ee) => {
          let re = J.replace(/^_xlfn\./i, "").toUpperCase();
          if (!r.some((ge) => ge.name === re)) return Z(J, ee);
          let oe = V.position;
          if (!oe) return f("#NAME?");
          let Y = j(oe.sheet).id,
            ne = `${Ue(oe.col - 1)}${oe.row}`,
            ae = ee.map((ge) => V.retrieveRef(ge)),
            le = Z2(ae);
          if (le) return le;
          let de = C6(e, re, ae, Y, ne);
          if (de.error) return f(de.error);
          let Se = de.value;
          return Se && !Array.isArray(Se) && typeof Se == "object"
            ? Se.label
            : Se;
        };
      }
      let G = (Z) => {
          let J = a.work() ?? i?.check();
          if (J) return f(J);
          let ee = Z(),
            re = a.value(ee) ?? i?.check();
          return re ? f(re) : ee;
        },
        H = V._callFunction.bind(V),
        K = V.utils._applyInfix.bind(V.utils),
        te = V.utils._applyPrefix.bind(V.utils),
        Q = V.utils._applyPostfix.bind(V.utils);
      if (
        ((V._callFunction = (Z, J) => G(() => H(Z, J))),
        (V.utils._applyInfix = (Z, J, ee) => G(() => K(Z, J, ee))),
        (V.utils._applyPrefix = (Z, J, ee) => G(() => te(Z, J, ee))),
        (V.utils._applyPostfix = (Z, J, ee) => G(() => Q(Z, J, ee))),
        n)
      ) {
        let Z = (Y, ne, ae, le) => {
            let de = V.position,
              Se = j(de.sheet).id,
              ge = `${Ue(de.col - 1)}${de.row}`;
            try {
              let _e = le();
              return (
                n.step({
                  kind: Y,
                  sheetId: Se,
                  address: ge,
                  operation: ne,
                  inputs: ae,
                  result: _e,
                }),
                _e
              );
            } catch (_e) {
              throw (
                n.step({
                  kind: Y,
                  sheetId: Se,
                  address: ge,
                  operation: ne,
                  inputs: ae,
                  result: dt(_e) ? _e : f("#ERROR!"),
                }),
                _e
              );
            }
          },
          J = V._callFunction.bind(V),
          ee = V.utils._applyInfix.bind(V.utils),
          re = V.utils._applyPrefix.bind(V.utils),
          oe = V.utils._applyPostfix.bind(V.utils);
        ((V._callFunction = (Y, ne) => Z("function", Y, ne, () => J(Y, ne))),
          (V.utils._applyInfix = (Y, ne, ae) =>
            Z("operator", ne, [Y.val, ae.val], () => ee(Y, ne, ae))),
          (V.utils._applyPrefix = (Y, ne, ae) =>
            Z("operator", Y.join(""), [ne], () => re(Y, ne, ae))),
          (V.utils._applyPostfix = (Y, ne, ae) =>
            Z("operator", ae, [Y], () => oe(Y, ne, ae))));
      }
      return V;
    },
    D = (V, G, H) =>
      c5(V, e, G, {
        supported: o,
        issue: f,
        isIssue: dt,
        allowName: n?.allowName,
        step: n
          ? (K, te, Q) =>
              n.step({
                kind: K,
                sheetId: G,
                address: H,
                operation: te,
                result: Q,
              })
          : undefined,
        diagnostic: (K, te) =>
          u.set(`${G}:${H}`, { sheetId: G, address: H, error: K, message: te }),
        parse: (K, te, Q) => {
          let Z = z(K, te, H, false),
            J = Z.replace(/"(?:[^"]|"")*"/g, '""');
          if (/[\[\]|]/.test(J) || /(?:https?:|file:|\\\\)/i.test(J))
            throw f("#BLOCKED!");
          if (
            Array.from(
              J.replace(/'(?:[^']|'')*'(?=\s*!)/g, "Sheet").matchAll(
                /([A-Za-z_][A-Za-z0-9_.]*)\s*\(/g,
              ),
            ).some((Y) => Y[1] !== Dm && !o.has(Y[1].toUpperCase()))
          )
            throw f("#NAME?");
          let ee = R++,
            re = E[ee] ?? (E[ee] = q()),
            oe = re._callFunction;
          re._callFunction = (Y, ne) =>
            Y === Dm
              ? ne.length === 1 &&
                Number.isInteger(ne[0]) &&
                Number(ne[0]) >= 0 &&
                Number(ne[0]) < Q.length
                ? Q[Number(ne[0])]
                : f("#NAME?")
              : oe(Y, ne);
          try {
            let Y = tt(H);
            return re.parse(
              Z,
              { sheet: v.get(te).name, row: Y.row + 1, col: Y.col + 1 },
              true,
            );
          } finally {
            ((re._callFunction = oe), R--);
          }
        },
      }),
    W = (V, G) => {
      let H = a.work() ?? i?.check();
      if (H) return f(H);
      if (!n) return $(V, G);
      if (!n.allowRead(V, G)) return f("#BLOCKED!");
      let K = $(V, G);
      return (
        n.step({
          kind: "read",
          sheetId: V,
          address: G,
          operation: "Read cell",
          result: K,
        }),
        K
      );
    },
    $ = (V, G) => {
      if (++x > Ne.maxCells) return f("#LIMIT!");
      let H, K;
      try {
        ({ row: H, col: K } = tt(G));
      } catch {
        return f("#REF!");
      }
      let te = v.get(V);
      if (!te || H >= te.rowCount || K >= te.colCount) return f("#REF!");
      G = `${Ue(K)}${H + 1}`;
      let Q = `${V}:${G}`,
        Z = L.values.get(Q);
      if (Z) return Z.type === "error" ? f(String(Z.value)) : Z.value;
      i?.read(Q);
      let J = i?.value(Q, w.size > 0 || _.size > 0);
      if (J) return J.type === "error" ? f(String(J.value)) : J.value;
      for (let ae of w) S.get(ae)?.add(Q);
      if (w.has(Q) || _.has(Q))
        return (i?.cycle([...w, ..._, Q]), f("#CYCLE!"));
      if (h.has(Q)) return h.get(Q);
      if (m.has(Q)) return m.get(Q).value;
      let ee = te.cells[G];
      if (!ee || ee.raw === null) {
        let ae = N.get(V) ?? [];
        if (ae.length > 1e3) return f("#LIMIT!");
        for (let le of ae) {
          if (le === G) continue;
          let de = tt(le);
          if (
            !(
              de.row > H ||
              de.col > K ||
              (H - de.row + 1) * (K - de.col + 1) > dr.maxSpillCells
            ) &&
            !(w.has(`${V}:${le}`) || _.has(`${V}:${le}`)) &&
            (W(V, le), m.has(Q))
          )
            return m.get(Q).value;
        }
        return ee?.raw ?? null;
      }
      if (typeof ee.raw != "string" || !ee.raw.startsWith("=") || ee.literal)
        return ee.raw;
      let re = ee.raw;
      i?.enter(Q);
      let Y = (() => {
          if (
            w.size + _.size >= dr.maxFormulaDepth ||
            re.length > dr.maxFormulaLength ||
            !Uu(re)
          )
            return f("#LIMIT!");
          let ae,
            le = A(re);
          _.add(Q);
          try {
            if (((ae = le ? re.slice(1) : z(re.slice(1), V, G)), !Uu(ae)))
              return f("#LIMIT!");
          } catch (_e) {
            return dt(_e)
              ? f(_e.error)
              : f(
                  _e instanceof Error && /^#/.test(_e.message)
                    ? _e.message
                    : "#NAME?",
                );
          } finally {
            _.delete(Q);
          }
          let de = ae.replace(/"(?:[^"]|"")*"/g, '""');
          if (!le && (/[\[\]|]/.test(de) || /(?:https?:|file:|\\\\)/i.test(de)))
            return f("#BLOCKED!");
          if (
            !le &&
            Array.from(
              de
                .replace(/'(?:[^']|'')*'(?=\s*!)/g, "Sheet")
                .matchAll(/([A-Za-z_][A-Za-z0-9_.]*)\s*\(/g),
            ).some((_e) => !o.has(_e[1].toUpperCase()))
          )
            return (
              u.set(Q, {
                sheetId: V,
                address: G,
                error: "#NAME?",
                message:
                  "The formula calls an unknown or unsupported custom function. Imported VBA, add-in and external custom functions cannot execute.",
              }),
              f("#NAME?")
            );
          let Se = w.size;
          (w.add(Q), S.set(Q, new Set()));
          let ge;
          try {
            let _e = O[Se] ?? (O[Se] = q()),
              Ge = le
                ? D(ae, V, G)
                : _e.parse(
                    ae,
                    { sheet: te.name, row: H + 1, col: K + 1 },
                    true,
                  );
            if (dt(Ge)) ge = f(Ge.error);
            else if (typeof Ge == "number")
              ge = Number.isFinite(Ge) ? Ge : f("#NUM!");
            else if (Ge === null) ge = 0;
            else if (typeof Ge == "boolean") ge = Ge;
            else if (typeof Ge == "string")
              ge = Ge.length <= 32767 ? Ge : f("#LIMIT!");
            else if (Array.isArray(Ge)) {
              let Le = Ge.every(Array.isArray) ? Ge : [Ge],
                Ae = Le[0]?.length ?? 0;
              if (!Le.length || !Ae || Le.some((je) => je.length !== Ae))
                ge = f("#CALC!");
              else if (
                Le.length * Ae > dr.maxSpillCells ||
                C + Le.length * Ae > Ne.maxCells
              )
                ge = f("#LIMIT!");
              else if (H + Le.length > te.rowCount || K + Ae > te.colCount)
                (b.set(Q, {
                  sheetId: V,
                  anchor: G,
                  reason: "bounds",
                  message:
                    "The array extends beyond the worksheet edge. Move the formula or reduce its output.",
                }),
                  (ge = f("#SPILL!")));
              else {
                let je = false,
                  Ie = false,
                  Te = {
                    firstRow: H,
                    lastRow: H + Le.length - 1,
                    firstCol: K,
                    lastCol: K + Ae - 1,
                  },
                  kt = (mr) => {
                    let $e = Qe(mr);
                    return (
                      $e.firstRow <= Te.lastRow &&
                      $e.lastRow >= Te.firstRow &&
                      $e.firstCol <= Te.lastCol &&
                      $e.lastCol >= Te.firstCol
                    );
                  },
                  Ve = te.merges?.find(kt),
                  et = te.tables?.find((mr) => kt(mr.range));
                if (Ve || et) {
                  je = true;
                  let mr = Ve ?? et.range;
                  b.set(Q, {
                    sheetId: V,
                    anchor: G,
                    reason: Ve ? "merged" : "table",
                    obstruction: mr,
                    message: Ve
                      ? `The array overlaps merged cells ${mr}. Unmerge those cells or move the formula.`
                      : `The array overlaps table ${et.name} (${mr}). Place the formula outside the table.`,
                  });
                }
                for (let mr = 0; mr < Le.length && !je; mr++)
                  for (let $e = 0; $e < Ae; $e++) {
                    let Je = `${Ue(K + $e)}${H + mr + 1}`,
                      gr = te.cells[Je];
                    if (S.get(Q)?.has(`${V}:${Je}`)) {
                      Ie = true;
                      break;
                    }
                    if (
                      Je !== G &&
                      ((gr && gr.raw !== null) || m.has(`${V}:${Je}`))
                    ) {
                      ((je = true),
                        b.set(Q, {
                          sheetId: V,
                          anchor: G,
                          reason: "occupied",
                          obstruction: Je,
                          message: `Cell ${Je} blocks the array. Move the formula or review and clear the obstruction; existing evidence was preserved.`,
                        }));
                      break;
                    }
                  }
                if (Ie) ge = f("#CYCLE!");
                else if (je) ge = f("#SPILL!");
                else {
                  let mr = ($e) =>
                    dt($e)
                      ? $e
                      : $e === null
                        ? 0
                        : typeof $e == "number"
                          ? Number.isFinite($e)
                            ? $e
                            : f("#NUM!")
                          : typeof $e == "boolean"
                            ? $e
                            : typeof $e == "string"
                              ? $e.length <= 32767
                                ? $e
                                : f("#LIMIT!")
                              : f("#VALUE!");
                  for (let $e = 0; $e < Le.length; $e++)
                    for (let Je = 0; Je < Ae; Je++) {
                      let gr = `${Ue(K + Je)}${H + $e + 1}`;
                      gr !== G &&
                        m.set(`${V}:${gr}`, {
                          value: mr(Le[$e][Je]),
                          anchor: G,
                        });
                    }
                  (g.set(Q, `${G}:${Ue(K + Ae - 1)}${H + Le.length}`),
                    (C += Le.length * Ae),
                    (ge = mr(Le[0][0])));
                }
              }
            } else ge = f("#VALUE!");
          } catch (_e) {
            let Ge =
              _e && typeof _e == "object" && "details" in _e
                ? _e.details
                : undefined;
            ge = dt(Ge) ? f(Ge.error) : dt(_e) ? f(_e.error) : f("#ERROR!");
          } finally {
            (w.delete(Q), S.delete(Q));
          }
          return ge;
        })(),
        ne = a.value(Y) ?? i?.check();
      return (
        ne && (Y = f(ne)),
        h.set(Q, Y),
        i?.leave(
          Q,
          dt(Y)
            ? { type: "error", value: Y.error }
            : {
                type:
                  Y === null
                    ? "blank"
                    : typeof Y == "string"
                      ? "text"
                      : typeof Y == "boolean"
                        ? "boolean"
                        : "number",
                value: Y,
              },
        ),
        n?.step({
          kind: "formula",
          sheetId: V,
          address: G,
          operation: ee.raw,
          result: Y,
        }),
        Y
      );
    },
    M = (V, G) => M.getTypedValue(V, G).value;
  M.getTypedValue = (V, G) => {
    if (i?.running?.() === false)
      try {
        let H = tt(G),
          K = v.get(V);
        if (!K || H.row >= K.rowCount || H.col >= K.colCount)
          return { type: "error", value: "#REF!" };
        let te = `${Ue(H.col)}${H.row + 1}`,
          Q = `${V}:${te}`,
          Z = i.value(Q, false);
        if (Z) return { ...Z };
        let J = h.get(Q) ?? m.get(Q)?.value ?? K.cells[te]?.raw ?? null;
        return dt(J)
          ? { type: "error", value: J.error }
          : {
              type:
                J === null
                  ? "blank"
                  : typeof J == "string"
                    ? "text"
                    : typeof J == "boolean"
                      ? "boolean"
                      : "number",
              value: J,
            };
      } catch {
        return { type: "error", value: "#REF!" };
      }
    a.begin();
    try {
      x = 0;
      let H = W(V, G);
      return dt(H)
        ? { type: "error", value: H.error }
        : {
            type:
              H === null
                ? "blank"
                : typeof H == "string"
                  ? "text"
                  : typeof H == "boolean"
                    ? "boolean"
                    : "number",
            value: H,
          };
    } finally {
      a.end();
    }
  };
  let B;
  return (
    (M.evaluateFormula = (V, G, H) => {
      (a.begin(), (x = 0), c.delete(`${V}:${G}`), u.delete(`${V}:${G}`));
      let K = `expression:${V}:${G}`;
      try {
        let te = v.get(V),
          Q = tt(G);
        if (!te || Q.row >= te.rowCount || Q.col >= te.colCount) return "#REF!";
        if (typeof H != "string" || !H.startsWith("=")) return "#VALUE!";
        if (H.length > dr.maxFormulaLength || !Uu(H)) return f("#LIMIT!").error;
        let Z = A(H),
          J = Z ? H.slice(1) : z(H.slice(1), V, G);
        if (!Uu(J)) return f("#LIMIT!").error;
        let ee = J.replace(/"(?:[^"]|"")*"/g, '""');
        if (!Z && (/[\[\]|]/.test(ee) || /(?:https?:|file:|\\\\)/i.test(ee)))
          return f("#BLOCKED!").error;
        if (
          !Z &&
          Array.from(
            ee
              .replace(/'(?:[^']|'')*'(?=\s*!)/g, "Sheet")
              .matchAll(/([A-Za-z_][A-Za-z0-9_.]*)\s*\(/g),
          ).some((Y) => !o.has(Y[1].toUpperCase()))
        )
          return "#NAME?";
        (w.add(K), S.set(K, new Set()));
        let re = Z
            ? D(J, V, G)
            : (B ??= q()).parse(
                J,
                { sheet: te.name, row: Q.row + 1, col: Q.col + 1 },
                true,
              ),
          oe = a.value(re) ?? i?.check();
        return oe
          ? f(oe).error
          : dt(re)
            ? f(re.error).error
            : Array.isArray(re)
              ? "#VALUE!"
              : re === null
                ? 0
                : typeof re == "number"
                  ? Number.isFinite(re)
                    ? re
                    : "#NUM!"
                  : typeof re == "boolean"
                    ? re
                    : typeof re == "string"
                      ? re.length <= 32767
                        ? re
                        : f("#LIMIT!").error
                      : "#VALUE!";
      } catch (te) {
        let Q =
          te && typeof te == "object" && "details" in te
            ? te.details
            : undefined;
        return dt(Q)
          ? f(Q.error).error
          : dt(te)
            ? f(te.error).error
            : te instanceof Error && /^#/.test(te.message)
              ? f(te.message).error
              : f("#ERROR!").error;
      } finally {
        (a.end(), w.delete(K), S.delete(K));
      }
    }),
    (M.getSpillRange = (V, G) => (M(V, G), g.get(`${V}:${G}`))),
    (M.getSpillDiagnostic = (V, G) => {
      M(V, G);
      let H = b.get(`${V}:${G}`);
      return H ? { ...H } : undefined;
    }),
    (M.getSpillAnchor = (V, G) => (M(V, G), m.get(`${V}:${G}`)?.anchor)),
    (M.getSpillEntries = (V) => {
      a.begin();
      try {
        let G = N.get(V) ?? [];
        if (G.length > 1e3)
          throw (
            f("#LIMIT!"),
            new Error("Spill discovery exceeds 1,000 possible array formulas.")
          );
        x = 0;
        for (let K of i?.running?.() === false ? [] : G)
          if ((W(V, K), x > Ne.maxCells))
            throw new Error(
              "Spill discovery exceeds 100,000 calculation reads.",
            );
        let H = `${V}:`;
        return [...m.entries()]
          .filter(([K]) => K.startsWith(H))
          .map(([K, { anchor }]) => ({
            address: K.slice(H.length),
            anchor,
          }));
      } finally {
        a.end();
      }
    }),
    (M.getResourceReport = () => a.report()),
    (M.getOperationalErrors = () => [...l]),
    (M.getDateContext = () => ({ ...s })),
    (M.getFinancialDiagnostics = () => structuredClone([...d.values()])),
    (M.peekSpill = (V, G) => {
      let H = `${V}:${G}`,
        K = b.get(H);
      return { range: g.get(H), ...(K ? { diagnostic: { ...K } } : {}) };
    }),
    (M.getNamedFormulaDiagnostics = () =>
      [...u.values()].map((V) => ({ ...V }))),
    (M.getReferenceDiagnostics = () => [...c.values()].map((V) => ({ ...V }))),
    M
  );
}
function yi(e, t = {}, n) {
  return !n &&
    !Object.keys(t).length &&
    (e.calculation || e.datePolicy || eR(e) || R6(e) || M6(e))
    ? tR(e).calculator
    : Hu(e, t, n);
}
var js = Object.freeze({
    mode: "automatic",
    iterative: false,
    maxIterations: 100,
    maxChange: 0.001,
    timeBudgetMs: 2500,
  }),
  Md = (e) =>
    !!e &&
    typeof e == "object" &&
    !Array.isArray(e) &&
    [Object.prototype, null].includes(Object.getPrototypeOf(e)),
  aW = () => {
    throw Error(
      "Invalid calculation settings. Choose Automatic or Manual, explicit iteration, 1\u20131000 iterations, convergence from 0 to 1, a 1\u201310000 ms time budget, 1\u2013128 MiB working data and 1000\u201310000000 work units.",
    );
  };
function $d(e) {
  return !Md(e) ||
    Object.keys(e).some(
      (t) => ![...Object.keys(js), "memoryBudgetMiB", "workBudget"].includes(t),
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
    ? aW()
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
function rR(e, t) {
  e.calculation !== undefined && (t.calculation = $d(e.calculation));
  let n = e.importedCalculationCaches;
  if (n === undefined) return;
  let i = () => {
    throw Error(
      "Invalid imported formula-cache evidence. Caches must remain separately labeled source evidence.",
    );
  };
  if (
    !Md(n) ||
    n.schemaVersion !== 1 ||
    typeof n.fileName != "string" ||
    !n.fileName ||
    n.fileName.length > 255 ||
    typeof n.sourceSha256 != "string" ||
    !/^[a-f0-9]{64}$/.test(n.sourceSha256) ||
    typeof n.importedAt != "string" ||
    !Number.isFinite(Date.parse(n.importedAt)) ||
    !Md(n.sourceSettings) ||
    !Array.isArray(n.entries) ||
    n.entries.length > 1e5
  )
    return i();
  let r = n.sourceSettings;
  if (
    typeof r.mode != "string" ||
    r.mode.length > 40 ||
    typeof r.iterative != "boolean" ||
    (r.maxIterations !== null &&
      (!Number.isSafeInteger(r.maxIterations) ||
        Number(r.maxIterations) < 1)) ||
    (r.fullPrecision !== undefined && typeof r.fullPrecision != "boolean") ||
    (r.maxChange !== null &&
      (typeof r.maxChange != "number" ||
        !Number.isFinite(r.maxChange) ||
        r.maxChange < 0))
  )
    return i();
  let o = (l) => {
      let c = /^([A-Z]+)([1-9]\d*)$/.exec(l);
      return (
        !!c &&
        Number(c[2]) <= 1048576 &&
        [...c[1]].reduce((u, d) => u * 26 + d.charCodeAt(0) - 64, 0) <= 16384
      );
    },
    a = new Set(),
    s = n.entries.map((l) => {
      if (
        !Md(l) ||
        typeof l.sheetId != "string" ||
        !l.sheetId ||
        l.sheetId.length > 100 ||
        typeof l.address != "string" ||
        !/^[A-Z]{1,3}[1-9]\d{0,6}$/.test(l.address) ||
        !o(l.address) ||
        typeof l.formula != "string" ||
        !l.formula.startsWith("=") ||
        l.formula.length > 8192
      )
        return i();
      let c = JSON.stringify([l.sheetId, l.address]);
      if (a.has(c)) return i();
      a.add(c);
      let u;
      if (l.cache !== undefined) {
        let d = l.cache;
        if (
          !Md(d) ||
          !["error", "blank", "number", "text", "boolean"].includes(
            String(d.type),
          ) ||
          (d.type === "blank" && d.value !== null) ||
          (d.type === "number" &&
            (typeof d.value != "number" || !Number.isFinite(d.value))) ||
          (d.type === "boolean" && typeof d.value != "boolean") ||
          (["text", "error"].includes(String(d.type)) &&
            (typeof d.value != "string" || d.value.length > 32767)) ||
          (d.type === "error" && !/^#[A-Z0-9_/?!]+$/.test(String(d.value)))
        )
          return i();
        u = { type: d.type, value: d.value };
      }
      return {
        sheetId: l.sheetId,
        address: l.address,
        formula: l.formula,
        ...(u ? { cache: u } : {}),
      };
    });
  t.importedCalculationCaches = {
    schemaVersion: 1,
    fileName: n.fileName,
    sourceSha256: n.sourceSha256,
    importedAt: n.importedAt,
    sourceSettings: {
      mode: r.mode,
      iterative: r.iterative,
      maxIterations: r.maxIterations,
      maxChange: r.maxChange,
      ...(r.fullPrecision !== undefined
        ? { fullPrecision: r.fullPrecision }
        : {}),
    },
    entries: s,
  };
}
var Pd = new WeakMap(),
  vc = new WeakMap(),
  nR = 2e5,
  Jm = { type: "error", value: "#PENDING!" },
  iR = (e) => ({ type: "error", value: e }),
  lW = new RegExp(
    `\\b(?:${[...kd, "LET", "LAMBDA", "INDIRECT", "OFFSET", "ISFORMULA", "TEXTBEFORE", "TEXTAFTER"].join("|")})\\s*\\(`,
    "i",
  ),
  oR = { test: (e) => Cs(e).length > 0 },
  Y2 = (e, t) =>
    !!e.definedNames?.length ||
    [...t.values()].some((n) => {
      let i = n.formula.replace(/"(?:[^"]|"")*"/g, '""');
      return (
        lW.test(i) ||
        /[{}#]/.test(i) ||
        /^=\s*[^()+*/]*:/.test(i) ||
        (/\b(?:IF|IFERROR|IFNA)\s*\(/i.test(i) && i.includes(":"))
      );
    }),
  sR = new WeakMap();
function cR(e) {
  if (!Hn(e)) return aR(e);
  let t = sR.get(e);
  return (t || ((t = aR(e)), sR.set(e, t)), t);
}
var Q2 = new WeakMap();
function cW(e) {
  let t = Hn(e.cells) ? Q2.get(e.cells)?.get(e.id) : undefined;
  if (t) return t;
  let n = [],
    i,
    r,
    o = () =>
      (i ??= Object.entries(e.cells).map(([l, c]) => [
        `${e.id}:${l}`,
        JSON.stringify([c.raw, !!c.literal]),
      ])),
    a = () => (r ??= JSON.stringify(o()));
  Hn(e.cells) || a();
  for (let l in e.cells) {
    let c = e.cells[l],
      u = `${e.id}:${l}`;
    !c.literal &&
      typeof c.raw == "string" &&
      c.raw.startsWith("=") &&
      n.push([u, { sheetId: e.id, address: l, formula: c.raw }]);
  }
  let s = {
    inputs: o,
    formulas: n,
    signature: a,
    volatile: [...new Set(n.flatMap(([, l]) => Cs(l.formula)))],
  };
  if (Hn(e.cells)) {
    let l = Q2.get(e.cells) ?? new Map();
    (l.set(e.id, s), Q2.set(e.cells, l));
  }
  return s;
}
function aR(e) {
  let t = e.sheets.map(cW),
    n,
    i = () => (n ??= new Map(t.flatMap((d) => d.inputs()))),
    r = new Map(t.flatMap((d) => d.formulas)),
    o = $d(e.calculation ?? js),
    a = ek(e),
    s,
    l,
    c = () =>
      (s ??= `[${JSON.stringify(e.revision)},${JSON.stringify(a)},[${t.map((d) => d.signature()).join(",")}]]`),
    u = () =>
      (l ??= `[${c()},${JSON.stringify(e.importedCalculationCaches ?? null)}]`);
  return (
    Hn(e) || u(),
    {
      inputs: i,
      formulas: r,
      settings: o,
      metadata: a,
      volatile: [
        ...new Set(
          t
            .flatMap((d) => d.volatile)
            .concat((e.definedNames ?? []).flatMap((d) => Cs(d.refersTo))),
        ),
      ].sort(),
      inputSignature: c,
      signature: u,
    }
  );
}
function uR(e, t, n = cR(t)) {
  return !!e && (Hn(t) || e.signature() === n.signature());
}
function ek(e) {
  let t = $d(e.calculation ?? js);
  return JSON.stringify([
    e.id,
    t,
    go(e.datePolicy),
    e.definedNames,
    e.workbookLinks,
    e.extensions,
    wp(e),
    e.sheets.map((i) => [
      i.id,
      i.name,
      i.rowCount,
      i.colCount,
      i.hiddenRows ?? [],
      i.hiddenCols ?? [],
      !!i.hidden,
      !!i.veryHidden,
      i.tables ?? [],
      i.merges ?? [],
      i.outlineGroups ?? [],
    ]),
  ]);
}
function lR(e, t) {
  if (
    e.revision !== t.revision ||
    ek(e) !== ek(t) ||
    (e.importedCalculationCaches !== t.importedCalculationCaches &&
      JSON.stringify(e.importedCalculationCaches) !==
        JSON.stringify(t.importedCalculationCaches))
  )
    return false;
  for (let n = 0; n < e.sheets.length; n++) {
    let i = e.sheets[n].cells,
      r = t.sheets[n].cells;
    if (i === r) continue;
    let o = Object.keys(i);
    if (o.length !== Object.keys(r).length) return false;
    for (let a of o) {
      let s = i[a],
        l = r[a];
      if (!l || s.raw !== l.raw || !!s.literal != !!l.literal) return false;
    }
  }
  return true;
}
function eR(e) {
  return Pd.has(e) || vc.has(e);
}
function dR(e, t) {
  if (e === t) return;
  let n = Pd.get(e);
  n && uR(n, e) && lR(e, t)
    ? Pd.set(t, n)
    : vc.has(e) && lR(e, t) && vc.set(t, vc.get(e));
}
function tR(e) {
  let t = cR(e),
    n = Pd.get(e);
  if (uR(n, e, t)) return { calculator: n.calculator, report: n.report };
  let i = uW(e, t, n ? undefined : vc.get(e), false);
  return (
    fR(i.report),
    Pd.set(e, i),
    vc.delete(e),
    { calculator: i.calculator, report: i.report }
  );
}
function fR(e) {
  !e ||
    typeof e != "object" ||
    Object.isFrozen(e) ||
    (Object.values(e).forEach(fR), Object.freeze(e));
}
function uW(e, t, n, i) {
  let r = xa(),
    { formulas, settings } = t,
    s = zu(e.datePolicy),
    l = xp(e);
  l.begin();
  let c = new Set(),
    u = new Map(),
    d = new Map(),
    f = new Map(),
    p = i
      ? "Explicit full recalculation"
      : n
        ? "Changed inputs and transitive dependents"
        : "First calculation of this snapshot";
  if (
    i ||
    !n ||
    t.metadata !== n.metadata ||
    Y2(e, formulas) ||
    Y2(e, n.formulas)
  )
    (formulas.forEach((j, q) => c.add(q)),
      n &&
        !i &&
        (p =
          t.metadata !== n.metadata
            ? "Calculation settings or structural metadata changed"
            : "Conservative recalculation for names, dynamic references or arrays"));
  else {
    let j = new Set(),
      q = t.inputs(),
      D = n.inputs();
    for (let M of new Set([...q.keys(), ...D.keys()]))
      q.get(M) !== D.get(M) && j.add(M);
    (formulas.forEach((M, B) => {
      (oR.test(M.formula) ||
        !n.values.has(B) ||
        ["#TIMEOUT!", "#LIMIT!"].includes(String(n.values.get(B)?.value))) &&
        j.add(B);
    }),
      n.report.cycles.flat().forEach((M) => j.add(M)));
    let W = new Map();
    n.dependencies.forEach((M, B) =>
      M.forEach((V) => {
        let G = W.get(V) ?? new Set();
        (G.add(B), W.set(V, G));
      }),
    );
    let $ = [...j];
    for (let M = 0; M < $.length; M++)
      for (let B of W.get($[M]) ?? []) j.has(B) || (j.add(B), $.push(B));
    j.forEach((M) => {
      formulas.has(M) && c.add(M);
    });
  }
  formulas.forEach((j, q) => {
    !c.has(q) &&
      n?.values.has(q) &&
      (d.set(q, n.values.get(q)), u.set(q, new Set(n.dependencies.get(q))));
  });
  let m = new Map(d),
    g = [],
    b = new Set(),
    w = true,
    _,
    S = [...u.values()].reduce((j, q) => j + q.size, 0),
    O = () => (
      w && !_ && (_ = l.check()),
      w && !_ && xa() - r > settings.timeBudgetMs && (_ = "#TIMEOUT!"),
      w ? _ : undefined
    ),
    E = (j, q) => {
      let D = [];
      return Hu(e, {}, undefined, {
        budget: l,
        running: () => w,
        dateContext: s,
        check: O,
        read($) {
          let M = D.at(-1);
          if (!w || !M) return;
          let B = u.get(M);
          B.has($) || (S >= nR ? (_ = "#LIMIT!") : (B.add($), S++));
        },
        value($, M) {
          return M && q?.has($) ? q.get($) : j.get($);
        },
        enter($) {
          (D.push($),
            (S -= u.get($)?.size ?? 0),
            u.set($, new Set()),
            f.set($, (f.get($) ?? 0) + 1));
        },
        leave($, M) {
          (d.set($, M), D.pop());
        },
        cycle($) {
          let M = $.at(-1),
            B = [...new Set($.slice($.indexOf(M), -1))];
          (B.forEach((V) => b.add(V)),
            B.length &&
              !g.some(
                (V) => V.length === B.length && V.every((G) => B.includes(G)),
              ) &&
              g.push(B));
        },
      });
    },
    R = settings.mode === "manual" && !i;
  R &&
    c.forEach((j) => {
      (d.set(j, Jm), m.set(j, Jm), u.set(j, new Set(n?.dependencies.get(j))));
    });
  let I = E(m),
    U = () => {
      for (let [j, q] of formulas)
        d.set(j, m.get(j) ?? I.getTypedValue(q.sheetId, q.address));
    };
  R || U();
  let P = {
    status: settings.iterative ? "not-needed" : "off",
    passes: 0,
    maxDelta: null,
    message: settings.iterative
      ? "No circular calculation was required."
      : "Iteration is disabled. Circular references return #CYCLE!.",
  };
  if (!R && b.size && settings.iterative) {
    let j = new Set(b),
      q = new Map();
    u.forEach((W, $) =>
      W.forEach((M) => {
        let B = q.get(M) ?? new Set();
        (B.add($), q.set(M, B));
      }),
    );
    let D = [...b];
    for (let W = 0; W < D.length; W++)
      for (let $ of q.get(D[W]) ?? []) j.has($) || (j.add($), D.push($));
    if (
      (d.forEach((W, $) => {
        j.has($) || m.set($, W);
      }),
      (P.status = "failed"),
      Y2(e, formulas) ||
        [...formulas.values()].some((W) => oR.test(W.formula)) ||
        b.size > 1e3)
    )
      P.message =
        "Iteration supports at most 1,000 deterministic scalar cycle cells. Names, arrays, dynamic references and volatile formulas must be removed from this calculation.";
    else {
      let W = new Map([...b].map(($) => [$, { type: "number", value: 0 }]));
      P.message = `No convergence within ${settings.maxIterations} simultaneous passes from zero.`;
      for (let $ = 0; $ < settings.maxIterations && !O(); $++) {
        I = E(m, W);
        let M = new Map(),
          B = 0,
          V = true;
        for (let G of b) {
          let H = formulas.get(G),
            K = I.getTypedValue(H.sheetId, H.address);
          if ((M.set(G, K), K.type !== "number" || !Number.isFinite(K.value))) {
            V = false;
            break;
          }
          B = Math.max(B, Math.abs(Number(K.value) - Number(W.get(G).value)));
        }
        if (
          ((P.passes = $ + 1), (P.maxDelta = Number.isFinite(B) ? B : null), !V)
        ) {
          P.message =
            "The circular calculation produced a nonnumeric value or calculation error.";
          break;
        }
        if (((W = M), B <= settings.maxChange)) {
          ((P.status = "converged"),
            (P.message = `Converged from zero in ${$ + 1} simultaneous passes; maximum change ${B}.`));
          break;
        }
      }
      P.status === "converged" && !_ && b.forEach(($) => m.set($, W.get($)));
    }
    ((P.status !== "converged" || _) &&
      b.forEach((W) => m.set(W, iR(_ ?? "#ITER!"))),
      (I = E(m)),
      U());
  }
  (O(),
    _ &&
      (c.forEach((j) => d.set(j, iR(_))),
      P.status === "converged" && (P.status = "failed"),
      b.size &&
        settings.iterative &&
        (P.message = `Calculation stopped with ${_}; no intermediate result is current.`)),
    (w = false),
    l.end());
  let T = (j, q) => T.getTypedValue(j, q).value;
  Object.assign(T, I);
  let A = (j, q) => `${j}:${q.replace(/\$/g, "").toUpperCase()}`,
    L = new Map();
  if (R && n)
    for (let j of e.sheets)
      for (let q of n.calculator.getSpillEntries(j.id))
        c.has(`${j.id}:${q.anchor}`) && L.set(`${j.id}:${q.address}`, q.anchor);
  T.getTypedValue = (j, q) => {
    let D = A(j, q);
    if (d.has(D)) return { ...d.get(D) };
    let W = e.sheets.find((M) => M.id === j)?.cells[
      q.replace(/\$/g, "").toUpperCase()
    ];
    if (
      W &&
      W.raw !== null &&
      (W.literal || typeof W.raw != "string" || !W.raw.startsWith("="))
    )
      return {
        type:
          typeof W.raw == "string"
            ? "text"
            : typeof W.raw == "boolean"
              ? "boolean"
              : "number",
        value: W.raw,
      };
    if (L.has(D)) return { ...Jm };
    let $ = I.getSpillAnchor(j, q);
    return $ && d.get(`${j}:${$}`)?.type === "error"
      ? { ...d.get(`${j}:${$}`) }
      : I.getTypedValue(j, q);
  };
  let y = I.getSpillEntries;
  ((T.getSpillEntries = (j) => [
    ...y(j),
    ...[...L]
      .filter(([q]) => q.startsWith(`${j}:`))
      .map(([q, D]) => ({ address: q.slice(j.length + 1), anchor: D })),
  ]),
    (T.getOperationalErrors = () => [
      ...new Set([
        ...I.getOperationalErrors(),
        ...[...d.values()]
          .filter(
            (j) =>
              j.type === "error" &&
              [
                "#PENDING!",
                "#TIMEOUT!",
                "#ITER!",
                "#LIMIT!",
                "#BLOCKED!",
                "#CYCLE!",
                "#ERROR!",
              ].includes(String(j.value)),
          )
          .map((j) => j.value),
      ]),
    ]));
  let v = (j, q) => [
      ...j.filter(
        (D) =>
          !c.has(`${D.sheetId}:${D.address}`) &&
          formulas.has(`${D.sheetId}:${D.address}`),
      ),
      ...q,
    ],
    k = n?.calculator.getReferenceDiagnostics() ?? [],
    x = n?.calculator.getNamedFormulaDiagnostics() ?? [],
    C = n?.calculator.getFinancialDiagnostics() ?? [];
  ((T.getFinancialDiagnostics = () =>
    structuredClone(v(C, I.getFinancialDiagnostics()))),
    (T.getReferenceDiagnostics = () => v(k, I.getReferenceDiagnostics())),
    (T.getNamedFormulaDiagnostics = () =>
      v(x, I.getNamedFormulaDiagnostics())));
  let N = new Map(n?.report.cells.map((j) => [`${j.sheetId}:${j.address}`, j])),
    F = new Map();
  for (let { sheetId, address, ...D } of T.getFinancialDiagnostics()) {
    let W = `${sheetId}:${address}`;
    F.set(W, [...(F.get(W) ?? []), D]);
  }
  let z = [...formulas].map(([j, q]) => {
    let D = d.get(j) ?? Jm,
      W = n?.values.get(j);
    return {
      ...q,
      ...(D.value !== "#PENDING!" && F.has(j) ? { financial: F.get(j) } : {}),
      result: { ...D },
      status:
        D.type === "error"
          ? D.value === "#PENDING!"
            ? "pending"
            : "error"
          : "current",
      ...(D.value === "#PENDING!" && W && W.value !== "#PENDING!"
        ? { lastKnown: { revision: n.report.revision, result: { ...W } } }
        : {}),
      ...(D.value === "#PENDING!" &&
      W?.value === "#PENDING!" &&
      N.get(j)?.lastKnown
        ? { lastKnown: N.get(j).lastKnown }
        : {}),
      dependencies: [...(u.get(j) ?? [])].sort(),
      evaluations: f.get(j) ?? 0,
    };
  });
  return {
    ...t,
    calculator: T,
    values: d,
    dependencies: u,
    report: {
      resources: l.report(),
      volatile: {
        functions: [
          ...new Set(
            [...formulas.values()]
              .flatMap((j) => Cs(j.formula))
              .concat((e.definedNames ?? []).flatMap((j) => Cs(j.refersTo))),
          ),
        ].sort(),
        policy: O6,
      },
      execution: {
        mode: "cooperative",
        state:
          _ === "#TIMEOUT!"
            ? "timeout"
            : _
              ? "failed"
              : R
                ? "pending"
                : "complete",
        message: _
          ? (l.report().reason ?? `Calculation stopped with ${_}.`)
          : "Synchronous calculation uses cooperative limits.",
      },
      dateContext: s,
      workbookId: e.id,
      revision: e.revision,
      run: (n?.report.run ?? 0) + 1,
      settings,
      reason: p,
      formulaCount: formulas.size,
      evaluated: z.filter((j) => j.evaluations > 0).length,
      reused: formulas.size - c.size,
      pending: z.filter((j) => j.status === "pending").length,
      elapsedMs: Math.max(0, xa() - r),
      dependencyEdges: [...u.values()].reduce((j, q) => j + q.size, 0),
      cycles: g,
      iteration: P,
      cells: z,
      limits: {
        dependencyEdges: nR,
        cooperativeTimeBudgetMs: settings.timeBudgetMs,
        hardInterrupt: false,
      },
    },
  };
}
var ss = Object.freeze({
  rows: 1e4,
  cells: 1e5,
  columns: 64,
  conditions: 128,
  values: 1e3,
  text: 1024,
  comparisonWork: 5e6,
});
var dW = [
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
  hR = (e) =>
    e === null ||
    typeof e == "boolean" ||
    (typeof e == "number" && Number.isFinite(e)) ||
    (typeof e == "string" && e.length <= 32767);
function pR(e, t, n) {
  if (
    !n ||
    typeof n.range != "string" ||
    !Array.isArray(n.columns) ||
    n.columns.length > ss.columns ||
    [n.caseSensitive, n.unique].some(
      (d) => d !== undefined && typeof d != "boolean",
    )
  )
    throw Error("Choose a header range and up to 64 column filters.");
  let i = mi(e, t),
    r = Qe(n.range),
    o = r.lastRow - r.firstRow,
    a = r.lastCol - r.firstCol + 1;
  if (r.lastRow >= i.rowCount || r.lastCol >= i.colCount || o < 1)
    throw Error(
      "Filter range needs one header row and at least one data row within the worksheet.",
    );
  if (o > ss.rows || (o + 1) * a > ss.cells || a > ss.columns)
    throw Error(
      "Filtering is limited to 10,000 data rows, 100,000 cells including headers, and 64 columns. Choose a smaller range.",
    );
  if (
    i.merges?.some((d) => {
      let f = Qe(d);
      return (
        f.firstRow <= r.lastRow &&
        f.lastRow >= r.firstRow &&
        f.firstCol <= r.lastCol &&
        f.lastCol >= r.firstCol
      );
    })
  )
    throw Error(
      "Unmerge cells in the filter range first; records cannot be split by filtering.",
    );
  let s = (d) => Number.isInteger(d) && d >= r.firstCol && d <= r.lastCol,
    l = (d) => {
      if (
        !d ||
        !s(d.column) ||
        !["value", "text", "number", "date"].includes(d.type) ||
        !dW.includes(d.operator) ||
        (!["blank", "notBlank"].includes(d.operator) && !hR(d.value))
      )
        throw Error("Invalid filter condition, column or value.");
      if (
        !["blank", "notBlank"].includes(d.operator) &&
        d.type === "text" &&
        typeof d.value != "string"
      )
        throw Error("Text criteria require a text value.");
      if (typeof d.value == "string" && d.value.length > ss.text)
        throw Error("Filter criteria text is limited to 1,024 characters.");
      if (
        [
          "contains",
          "notContains",
          "startsWith",
          "endsWith",
          "wildcard",
        ].includes(d.operator) &&
        d.type !== "text"
      )
        throw Error("Text matching requires a Text condition.");
      if (["gt", "gte", "lt", "lte"].includes(d.operator) && d.type === "value")
        throw Error("Choose Number, Date or Text for ordered comparisons.");
      if (
        !["blank", "notBlank"].includes(d.operator) &&
        ["number", "date"].includes(d.type) &&
        fW(d.value, d.type, e) === null
      )
        throw Error(
          `Invalid ${d.type} filter value. Use a finite number or a date in the workbook's date policy.`,
        );
    },
    c = 0,
    u = new Set();
  for (let d of n.columns) {
    if (
      !d ||
      !s(d.column) ||
      u.has(d.column) ||
      (d.join !== undefined && !["and", "or"].includes(d.join))
    )
      throw Error("Filter columns must be valid and distinct.");
    if (
      (u.add(d.column),
      d.values !== undefined &&
        (!Array.isArray(d.values) ||
          d.values.length > ss.values ||
          !d.values.every(hR)))
    )
      throw Error("Value filters allow up to 1,000 finite typed values.");
    if (
      d.conditions !== undefined &&
      (!Array.isArray(d.conditions) || d.conditions.length > 2)
    )
      throw Error("Each column supports up to two custom conditions.");
    for (let f of d.conditions ?? []) {
      if (f.column !== d.column)
        throw Error("A column condition must use its own column.");
      (l(f), c++);
    }
  }
  if (
    n.advanced !== undefined &&
    (!Array.isArray(n.advanced) || n.advanced.length > ss.conditions)
  )
    throw Error("Invalid advanced filter groups.");
  for (let d of n.advanced ?? []) {
    if (!Array.isArray(d) || !d.length || d.length > ss.conditions)
      throw Error(
        "Each advanced OR group must contain at least one AND condition.",
      );
    for (let f of d) (l(f), c++);
  }
  if (c > ss.conditions)
    throw Error("At most 128 combined filter conditions are allowed.");
  return structuredClone({
    ...n,
    range: `${Ue(r.firstCol)}${r.firstRow + 1}:${Ue(r.lastCol)}${r.lastRow + 1}`,
  });
}
function fW(e, t, n) {
  if (t === "number") return typeof e == "number" ? e : null;
  if (t === "date") {
    let i = go(n.datePolicy),
      r =
        typeof e == "number" ? e : typeof e == "string" ? Hi(e, i) : undefined;
    return r !== undefined && r >= 0 && r < Xi(i.dateSystem) + 1 ? r : null;
  }
  return typeof e == "string" ? e : null;
}
var tk = Object.freeze({ views: 20, name: 80, characters: 512e3 });
function wc(e, t) {
  if (
    !e ||
    typeof e != "object" ||
    Array.isArray(e) ||
    ![Object.prototype, null].includes(Object.getPrototypeOf(e)) ||
    Object.keys(e).some((n) => !t.includes(n))
  )
    throw Error("Invalid saved filter view metadata.");
}
function Dd(e, t) {
  return (
    typeof e == "string" &&
    !!e.trim() &&
    e.length <= t &&
    !/[\u0000-\u001f\u007f]/.test(e)
  );
}
function mR(e) {
  return e.normalize("NFC").trim().toLocaleLowerCase("en-US");
}
function pW(e) {
  if (!Dd(e, tk.name))
    throw Error("Use a nonempty filter view name of up to 80 characters.");
  return e.normalize("NFC").trim();
}
function gR(e, t) {
  if (e === undefined) return;
  if (
    !Array.isArray(e) ||
    e.length > tk.views ||
    JSON.stringify(e).length > tk.characters
  )
    throw Error(
      "Saved filter views are limited to 20 views and 512,000 metadata characters per workbook.",
    );
  let n = new Set(),
    i = new Set();
  t.savedFilterViews = e.map((r) => {
    if (
      (wc(r, [
        "id",
        "name",
        "sheetId",
        "sheetName",
        "sourceWorkbookId",
        "sourceRevision",
        "savedAt",
        "datePolicy",
        "definition",
        "population",
      ]),
      !Dd(r.id, 100) ||
        !Dd(r.sheetId, 100) ||
        !Dd(r.sheetName, 31) ||
        !Dd(r.sourceWorkbookId, 100) ||
        !Number.isSafeInteger(r.sourceRevision) ||
        r.sourceRevision < 0 ||
        r.sourceRevision > Number.MAX_SAFE_INTEGER - 1 ||
        typeof r.savedAt != "string" ||
        !Number.isFinite(Date.parse(r.savedAt)) ||
        new Date(r.savedAt).toISOString() !== r.savedAt)
    )
      throw Error("Invalid saved filter view identity, revision or timestamp.");
    let o = pW(r.name);
    if (n.has(r.id) || i.has(mR(o)))
      throw Error("Saved filter view names and IDs must be unique.");
    (n.add(r.id),
      i.add(mR(o)),
      wc(r.datePolicy, [
        "calendar",
        "dateSystem",
        "textDateOrder",
        "timeZone",
      ]));
    let a = go(r.datePolicy);
    wc(r.definition, [
      "range",
      "columns",
      "advanced",
      "caseSensitive",
      "unique",
    ]);
    let s = (d) => wc(d, ["column", "type", "operator", "value"]);
    if (!Array.isArray(r.definition.columns))
      throw Error("Invalid saved filter columns.");
    for (let d of r.definition.columns)
      if (
        (wc(d, ["column", "values", "conditions", "join"]),
        d.conditions !== undefined)
      ) {
        if (!Array.isArray(d.conditions))
          throw Error("Invalid saved filter conditions.");
        d.conditions.forEach(s);
      }
    if (r.definition.advanced !== undefined) {
      if (!Array.isArray(r.definition.advanced))
        throw Error("Invalid saved advanced criteria.");
      for (let d of r.definition.advanced) {
        if (!Array.isArray(d)) throw Error("Invalid saved advanced criteria.");
        d.forEach(s);
      }
    }
    let l = {
        ...t,
        datePolicy: a,
        sheets: [
          {
            id: r.sheetId,
            name: r.sheetName,
            rowCount: Ne.maxRows,
            colCount: Ne.maxCols,
            cells: {},
          },
        ],
      },
      c = pR(l, r.sheetId, r.definition);
    wc(r.population, [
      "totalRows",
      "matchedRows",
      "visibleRows",
      "hiddenMatchingRows",
    ]);
    let u = r.population;
    if (
      Object.keys(u).length !== 4 ||
      Object.values(u).some((d) => !Number.isInteger(d) || Number(d) < 0) ||
      u.totalRows !== Qe(c.range).lastRow - Qe(c.range).firstRow ||
      u.totalRows > 1e4 ||
      u.totalRows < 1 ||
      u.matchedRows > u.totalRows ||
      u.visibleRows > u.matchedRows ||
      u.hiddenMatchingRows !== u.matchedRows - u.visibleRows
    )
      throw Error("Invalid saved population counts.");
    return {
      id: r.id,
      name: o,
      sheetId: r.sheetId,
      sheetName: r.sheetName,
      sourceWorkbookId: r.sourceWorkbookId,
      sourceRevision: r.sourceRevision,
      savedAt: r.savedAt,
      datePolicy: a,
      definition: c,
      population: { ...u },
    };
  });
}
var bR = "ledgerglass-consolidation-v1",
  yR = Object.freeze({
    sources: 8,
    categories: 3,
    contributions: 2e3,
    outputRows: 4e3,
    outputCells: 4e4,
    groups: 1e3,
    metadataCharacters: 8e6,
  }),
  ai = X.string()
    .min(1)
    .max(128)
    .refine((e) => !!e.trim() && !/[\u0000-\u001f\u007f]/.test(e)),
  Ui = X.string().min(1).max(100),
  nk = X.string().regex(/^[a-f0-9]{64}$/),
  vR = X.union([
    X.string().min(1).max(32767),
    X.number().finite(),
    X.boolean(),
  ]),
  rk = X.discriminatedUnion("kind", [
    X.strictObject({ kind: X.literal("constant"), value: ai }),
    X.strictObject({ kind: X.literal("field"), field: ai }),
  ]),
  mW = X.strictObject({
    id: Ui,
    workbookId: Ui,
    sheetId: Ui,
    range: X.string().min(1).max(50),
    tableId: Ui.optional(),
    categories: X.array(ai).min(1).max(3),
    amountField: ai,
    period: rk,
    currency: rk,
    unit: rk,
    duplicatePolicy: X.enum(["reject", "sum"]),
  }),
  gW = X.strictObject({
    engine: X.literal(bR),
    sources: X.array(mW).min(1).max(8),
    categoryLabels: X.array(ai).min(1).max(3),
    period: ai,
    currency: ai,
    unit: ai,
    linkedAmounts: X.boolean(),
    collapseDetails: X.boolean(),
  }).superRefine((e, t) => {
    let n = [
      "Row type",
      ...e.categoryLabels,
      "Period",
      "Currency",
      "Unit",
      "Amount",
      "Source",
      "Link state",
    ];
    (new Set(n.map((i) => i.toLowerCase())).size !== n.length ||
      new Set(e.sources.map((i) => i.id)).size !== e.sources.length ||
      e.sources.some(
        (i) =>
          i.categories.length !== e.categoryLabels.length ||
          new Set(i.categories).size !== i.categories.length,
      )) &&
      t.addIssue({
        code: "custom",
        message:
          "Use unique source IDs, distinct category fields, and distinct output headers with matching category counts.",
      });
  }),
  bW = X.strictObject({
    sourceId: Ui,
    workbookId: Ui,
    workbookName: X.string().min(1).max(200),
    sheetId: Ui,
    sheetName: X.string().min(1).max(31),
    row: X.number().int().min(2).max(1048576),
    address: X.string().regex(/^[A-Z]{1,3}[1-9]\d{0,6}$/),
    locator: X.string().max(300),
    categories: X.array(vR).min(1).max(3),
    amount: X.number().finite(),
    period: ai,
    currency: ai,
    unit: ai,
    link: X.enum(["live", "snapshot", "approved snapshot"]),
  }),
  yW = X.strictObject({
    sourceId: Ui,
    row: X.number().int().positive().optional(),
    code: X.string().max(40),
    severity: X.enum(["error", "warning"]),
    message: X.string().max(4e3),
  }),
  vW = X.strictObject({
    sourceId: Ui,
    workbookId: Ui,
    workbookName: X.string().max(200),
    sheetId: Ui,
    sheetName: X.string().max(31),
    range: X.string().max(50),
    fingerprint: nk,
    records: X.number().int().min(0).max(500),
    acceptedRecords: X.number().int().min(0).max(500),
    amount: X.number().finite(),
    periods: X.array(ai).max(500),
    currencies: X.array(ai).max(500),
    units: X.array(ai).max(500),
  }),
  wW = X.strictObject({
    categories: X.array(vR).min(1).max(3),
    level: X.number().int().min(1).max(3),
    count: X.number().int().min(1).max(2e3),
    amount: X.number().finite(),
  }),
  xW = X.strictObject({
    sources: X.array(vW).max(8),
    contributions: X.array(bW).max(2e3),
    subtotals: X.array(wW).max(1e3),
    exceptions: X.array(yW).max(8e3),
    grandTotal: X.number().finite().nullable(),
    sourceTotal: X.number().finite().nullable(),
    difference: X.number().finite().nullable(),
    complete: X.boolean(),
  }),
  kW = X.strictObject({
    engine: X.literal(bR),
    definition: gW,
    review: xW,
    origin: X.strictObject({
      workbookId: Ui,
      revision: X.number().int().nonnegative(),
      at: X.string().datetime(),
    }),
    evidenceHash: nk,
    outputRange: X.string().max(50),
    outputHash: nk,
  });
var _W = (e) => JSON.stringify(e.map((t) => [typeof t, t]));
function Ym(e) {
  let t = 0,
    n = 0;
  for (let i of e) {
    let r = i - n,
      o = t + r;
    if (((n = o - t - r), (t = o), !Number.isFinite(t) || !Number.isFinite(n)))
      throw Error("Consolidation total exceeds finite numeric limits.");
  }
  return Object.is(t, -0) ? 0 : t;
}
function EW(e) {
  let t = new Map();
  for (let n of e)
    for (let i = 1; i <= n.categories.length; i++) {
      let r = _W(n.categories.slice(0, i)),
        o = t.get(r) ?? [];
      (o.push(n), t.set(r, o));
    }
  if (t.size > yR.groups)
    throw Error(
      "Consolidation exceeds 1,000 subtotal groups. Narrow the sources; no rows were truncated.",
    );
  return [...t].map(([n, i]) => {
    let r = JSON.parse(n).length;
    return {
      categories: i[0].categories.slice(0, r),
      level: r,
      count: i.length,
      amount: Ym(i.map((o) => o.amount)),
    };
  });
}
function wR(e, t) {
  let n = 0;
  for (let i of t.sheets) {
    let r = e.sheets?.find((c) => c.id === i.id)?.consolidation;
    if (r === undefined) continue;
    if (((n += JSON.stringify(r).length), n > yR.metadataCharacters))
      throw Error(
        "Consolidation evidence exceeds 8,000,000 characters. Retain native backups before removing obsolete reports.",
      );
    let o = kW.safeParse(r);
    if (!o.success)
      throw Error(
        `Invalid consolidation evidence: ${o.error.issues[0]?.message}`,
      );
    let a = o.data,
      s = a.review,
      l = a.definition.sources.map((c) => c.id);
    if (
      !s.complete ||
      s.exceptions.some((c) => c.severity === "error") ||
      !s.contributions.length ||
      s.sources.length !== l.length ||
      new Set(s.sources.map((c) => c.sourceId)).size !== l.length ||
      s.sources.some((c) => !l.includes(c.sourceId)) ||
      s.contributions.some(
        (c) =>
          !l.includes(c.sourceId) ||
          c.categories.length !== a.definition.categoryLabels.length ||
          c.period !== a.definition.period ||
          c.currency !== a.definition.currency ||
          c.unit !== a.definition.unit,
      ) ||
      s.grandTotal !== Ym(s.contributions.map((c) => c.amount)) ||
      s.sourceTotal !== Ym(s.sources.map((c) => c.amount)) ||
      s.difference !== s.grandTotal - s.sourceTotal ||
      Jr(s.subtotals) !== Jr(EW(s.contributions)) ||
      a.evidenceHash !==
        Jr({ definition: a.definition, review: s, origin: a.origin })
    )
      throw Error(
        "Consolidation evidence does not reconcile to its sources, categories or retained fingerprint.",
      );
    for (let c of s.sources) {
      let u = s.contributions.filter((d) => d.sourceId === c.sourceId);
      if (
        c.records !== u.length ||
        c.acceptedRecords !== u.length ||
        c.amount !== Ym(u.map((d) => d.amount)) ||
        u.some((d) => d.workbookId !== c.workbookId || d.sheetId !== c.sheetId)
      )
        throw Error("Consolidation source evidence does not reconcile.");
    }
    i.consolidation = a;
  }
}
var kR = "ledgerglass-sources-v1",
  SW = ["database", "odbc", "odata", "api", "erp", "file-folder"],
  Wa = Object.freeze({
    sources: 20,
    columns: 50,
    rows: 5e3,
    cells: 1e5,
    responseBytes: 8 * 1024 * 1024,
    timeoutMs: 3e4,
    metadataCharacters: 2e5,
  }),
  Qm = X.string().regex(/^[A-Za-z0-9][A-Za-z0-9._-]{0,99}$/),
  eg = X.string()
    .min(1)
    .max(128)
    .refine((e) => e === e.trim() && !/[\u0000-\u001f\u007f]/.test(e)),
  Ld = X.string().datetime({ offset: true }),
  xR = X.string().regex(/^[a-f0-9]{64}$/),
  CW = X.strictObject({
    name: eg,
    type: X.enum(["text", "number", "boolean", "date"]),
    nullable: X.boolean(),
  }),
  ik = X.strictObject({
    id: Qm,
    label: eg,
    kind: X.enum(SW),
    schemaVersion: Qm,
    columns: X.array(CW).min(1).max(Wa.columns),
    controlFields: X.array(eg).max(Wa.columns),
  }).superRefine((e, t) => {
    (new Set(e.columns.map((n) => n.name.toLowerCase())).size !==
      e.columns.length ||
      new Set(e.controlFields).size !== e.controlFields.length ||
      e.controlFields.some(
        (n) => !e.columns.some((i) => i.name === n && i.type === "number"),
      )) &&
      t.addIssue({
        code: "custom",
        message: "Use unique columns and declared numeric control fields.",
      });
  }),
  Zte = X.strictObject({ source: ik, expiresAt: Ld }),
  TW = X.union([
    X.string().max(32767),
    X.number().finite(),
    X.boolean(),
    X.null(),
  ]),
  Kte = X.strictObject({
    protocol: X.literal(kR),
    source: ik,
    sourceTimestamp: Ld,
    version: Qm,
    rows: X.array(X.array(TW).max(Wa.columns)).max(Wa.rows),
    complete: X.literal(true),
  }),
  AW = [
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
  IW = {
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
  },
  Ud = class extends Error {
    constructor(n) {
      super(IW[n]);
      this.code = n;
      this.name = "ExternalSourceError";
    }
  };
var RW = X.strictObject({
    field: eg,
    sum: X.number().finite(),
    numeric: X.number().int().min(0).max(5e3),
    blanks: X.number().int().min(0).max(5e3),
  }),
  NW = X.strictObject({
    protocol: X.literal(kR),
    source: ik,
    lastSuccess: X.strictObject({
      at: Ld,
      sourceTimestamp: Ld,
      version: Qm,
      rows: X.number().int().min(0).max(5e3),
      controls: X.array(RW).max(50),
      dataHash: xR,
    }),
    lastAttempt: X.strictObject({
      at: Ld,
      outcome: X.enum(["succeeded", "failed", "cancelled"]),
      code: X.enum(AW).optional(),
    }),
    outputHash: xR,
  });
function _R(e, t) {
  let n = 0,
    i = 0;
  for (let r of t.sheets) {
    let o = e.sheets.find((l) => l.id === r.id)?.externalConnection;
    if (o === undefined) continue;
    if (
      ++i > Wa.sources ||
      (n += JSON.stringify(o).length) > Wa.metadataCharacters
    )
      throw new Ud("limit");
    let a = NW.safeParse(o);
    if (!a.success) throw new Ud("invalid");
    let s = a.data;
    if (
      s.lastSuccess.controls.length !== s.source.controlFields.length ||
      s.lastSuccess.controls.some(
        (l, c) =>
          l.field !== s.source.controlFields[c] ||
          l.numeric + l.blanks !== s.lastSuccess.rows,
      ) ||
      (s.lastSuccess.rows + 1) * s.source.columns.length > Wa.cells ||
      Date.parse(s.lastAttempt.at) < Date.parse(s.lastSuccess.at) ||
      (s.lastAttempt.outcome === "succeeded") == !!s.lastAttempt.code
    )
      throw new Ud("invalid");
    r.externalConnection = s;
  }
}
var OW = "ledgerglass-analysis-v1",
  ok = Object.freeze({
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
  Io = X.number().finite(),
  xc = Io.min(-1e9).max(1e9),
  tg = X.strictObject({
    sheetId: X.string().min(1).max(100),
    address: X.string().regex(/^[A-Z]{1,3}[1-9]\d{0,6}$/),
  }),
  ER = X.string()
    .trim()
    .min(1)
    .max(31)
    .refine(
      (e) => !/[\\/?*\[\]:\u0000-\u001f]/.test(e),
      "Invalid output worksheet name",
    ),
  MW = X.strictObject({
    kind: X.literal("solver"),
    outputName: ER,
    objective: tg,
    direction: X.enum(["maximize", "minimize", "target"]),
    target: xc.optional(),
    variables: X.array(
      X.strictObject({
        cell: tg,
        lower: xc,
        upper: xc,
        domain: X.enum(["continuous", "integer", "binary"]),
      }),
    )
      .min(1)
      .max(6),
    constraints: X.array(
      X.strictObject({
        left: tg,
        operator: X.enum(["<=", ">=", "="]),
        right: X.union([xc, tg]),
      }),
    ).max(16),
    tolerance: Io.min(1e-10).max(1e-5).default(1e-8),
    maxCandidates: X.number().int().min(1).max(5e4).default(5e4),
  }),
  $W = X.strictObject({
    kind: X.literal("statistics"),
    outputName: ER,
    sheetId: X.string().min(1).max(100),
    range: X.string().min(1).max(50),
    method: X.enum([
      "descriptive",
      "correlation",
      "covariance",
      "regression",
      "histogram",
      "distribution",
      "forecast",
    ]),
    missing: X.enum(["reject", "exclude"]).default("reject"),
    bins: X.array(xc).min(1).max(50).optional(),
    distribution: X.enum(["normal", "binomial", "poisson"]).optional(),
    mean: xc.optional(),
    standardDeviation: Io.positive().max(1e9).optional(),
    trials: X.number().int().min(0).max(1e3).optional(),
    probability: Io.min(0).max(1).optional(),
    lambda: Io.positive().max(1e3).optional(),
    forecast: X.enum(["linear", "moving-average", "exponential"]).optional(),
    horizon: X.number().int().min(1).max(24).optional(),
    window: X.number().int().min(2).max(100).optional(),
    alpha: Io.positive().max(1).optional(),
  }),
  SR = X.discriminatedUnion("kind", [MW, $W]),
  PW = X.union([X.string().max(32767), Io, X.boolean(), X.null()]),
  DW = X.strictObject({
    status: X.enum([
      "complete",
      "optimal",
      "infeasible",
      "work-limit",
      "precision-limit",
    ]),
    success: X.boolean(),
    message: X.string().max(2e3),
    algorithm: X.string().max(300),
    assumptions: X.array(X.string().max(1e3)).max(30),
    exclusions: X.array(
      X.strictObject({
        row: X.number().int().positive(),
        reason: X.string().max(200),
        locator: X.string().max(300),
      }),
    ).max(500),
    included: X.number().int().nonnegative(),
    inspected: X.number().int().nonnegative(),
    iterations: X.number().int().nonnegative(),
    rows: X.array(X.array(PW).max(16)).max(3e3),
    solution: X.array(Io).max(6).optional(),
    objective: Io.optional(),
    baseline: X.array(Io).max(6).optional(),
  }),
  sk = X.string().regex(/^[a-f0-9]{64}$/),
  LW = X.strictObject({
    engine: X.literal(OW),
    request: SR,
    result: DW,
    definitionHash: sk,
    sourceHash: sk,
    outputHash: sk,
    outputRange: X.string().max(50),
    createdAt: X.string().datetime(),
    sourceRevision: X.number().int().nonnegative(),
    includeHidden: X.boolean(),
    applied: X.boolean(),
  });
function UW(e) {
  let t = SR.safeParse(e);
  if (!t.success)
    throw Error(
      `Invalid analysis request: ${t.error.issues[0]?.path.join(".")}: ${t.error.issues[0]?.message}`,
    );
  let n = t.data;
  if (n.kind === "solver") {
    if (n.direction === "target" && n.target === undefined)
      throw Error("Enter a finite target value.");
    if (
      new Set(n.variables.map((i) => JSON.stringify(i.cell))).size !==
      n.variables.length
    )
      throw Error("Changing cells must be unique.");
    if (
      n.variables.some(
        (i) =>
          i.lower > i.upper ||
          (i.domain === "binary" && (i.lower < 0 || i.upper > 1)),
      )
    )
      throw Error(
        "Variable bounds must be ordered; binary bounds must lie within zero and one.",
      );
    if (
      n.variables.some((i) => i.domain === "continuous") &&
      n.variables.some((i) => i.domain !== "continuous")
    )
      throw Error(
        "Mixed integer and continuous models are outside this selected Solver subset.",
      );
  }
  return n;
}
function CR(e, t) {
  let n = e,
    i = 0;
  for (let r of t.sheets) {
    let o = n.sheets.find((u) => u.id === r.id)?.analysisReport;
    if (o === undefined) continue;
    if (((i += JSON.stringify(o).length), i > ok.metadataCharacters))
      throw Error("Analysis metadata exceeds 2,000,000 characters.");
    let a = LW.safeParse(o);
    if (!a.success)
      throw Error(`Invalid analysis report: ${a.error.issues[0]?.message}`);
    let s = a.data;
    if (
      ((s.request = UW(s.request)),
      s.definitionHash !== Jr(s.request) ||
        !s.result.success ||
        !["complete", "optimal"].includes(s.result.status) ||
        s.result.rows.reduce((u, d) => u + d.length, 0) > ok.outputCells)
    )
      throw Error("Analysis report metadata is inconsistent.");
    let l = /^([A-Z]{1,3})([1-9]\d{0,6}):([A-Z]{1,3})([1-9]\d{0,6})$/.exec(
        s.outputRange,
      ),
      c = (u) => [...u].reduce((d, f) => d * 26 + f.charCodeAt(0) - 64, 0);
    if (
      !l ||
      c(l[3]) > 16384 ||
      Number(l[4]) > 1048576 ||
      c(l[3]) < c(l[1]) ||
      Number(l[4]) < Number(l[2]) ||
      (c(l[3]) - c(l[1]) + 1) * (Number(l[4]) - Number(l[2]) + 1) >
        ok.outputCells
    )
      throw Error("Invalid or oversized analysis output range.");
    if (
      (s.request.kind === "solver" &&
        (s.result.status !== "optimal" ||
          s.result.solution?.length !== s.request.variables.length ||
          s.result.baseline?.length !== s.request.variables.length ||
          s.result.objective === undefined)) ||
      (s.request.kind === "statistics" &&
        (s.result.status !== "complete" || s.applied))
    )
      throw Error("Analysis result does not match its request.");
    r.analysisReport = s;
  }
}
var FW = "ledgerglass-report-output-v1",
  zW = Object.freeze({
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
  Si = (e) => X.string().max(e),
  Ws = X.string().min(1).max(200),
  rg = X.string().regex(/^[a-f0-9]{64}$/),
  IR = X.string().datetime();
function RR(e) {
  if (
    !/^\$?[A-Za-z]{1,3}\$?[1-9]\d{0,6}(?::\$?[A-Za-z]{1,3}\$?[1-9]\d{0,6})?$/.test(
      e,
    )
  )
    throw Error("Choose one explicit report range.");
  let t = e.replaceAll("$", "").toUpperCase().split(":"),
    n = Qe(e);
  if (
    t[0] !== `${Ue(n.firstCol)}${n.firstRow + 1}` ||
    (t[1] && t[1] !== `${Ue(n.lastCol)}${n.lastRow + 1}`)
  )
    throw Error("Report ranges cannot be reversed.");
  if (
    n.lastCol - n.firstCol + 1 > 20 ||
    n.lastRow - n.firstRow + 1 > 500 ||
    (n.lastCol - n.firstCol + 1) * (n.lastRow - n.firstRow + 1) > 2e3
  )
    throw Error(
      "Reports support at most 500 rows, 20 columns and 2,000 cells.",
    );
  return `${Ue(n.firstCol)}${n.firstRow + 1}:${Ue(n.lastCol)}${n.lastRow + 1}`;
}
var TR = X.object({
    column: X.number().int().min(0).max(16383),
    type: X.enum(["value", "text", "number", "date"]),
    operator: X.enum([
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
    value: X.union([
      X.string().max(1024),
      X.number().finite(),
      X.boolean(),
      X.null(),
    ]).optional(),
  }).strict(),
  BW = X.object({
    range: Si(40),
    columns: X.array(
      X.object({
        column: X.number().int().min(0).max(16383),
        values: X.array(
          X.union([Si(32767), X.number().finite(), X.boolean(), X.null()]),
        )
          .max(1e3)
          .optional(),
        conditions: X.array(TR).max(128).optional(),
        join: X.enum(["and", "or"]).optional(),
      }).strict(),
    ).max(20),
    advanced: X.array(X.array(TR).max(128)).max(128).optional(),
    caseSensitive: X.boolean().optional(),
    unique: X.boolean().optional(),
  }).strict(),
  NR = X.object({
    name: X.string().trim().min(1).max(100),
    kind: X.enum(["range", "object"]),
    sheetId: Ws,
    range: Si(40).transform(RR),
    objectId: Ws.optional(),
    mode: X.enum(["snapshot", "linked"]),
    population: X.enum(["full", "visible"]),
    filter: BW.optional(),
  }).strict(),
  jW = X.object({
    address: X.string().regex(/^[A-Z]{1,3}[1-9]\d{0,6}$/),
    value: X.union([Si(32767), X.number().finite(), X.boolean(), X.null()]),
    text: Si(32767),
    style: X.unknown().optional(),
  }).strict(),
  WW = X.object({
    engine: X.literal(FW),
    id: Ws,
    request: NR,
    workbookId: Ws,
    workbookName: Si(200),
    revision: X.number().int().nonnegative(),
    sheetName: Si(31),
    calculatedAt: IR,
    sourceHash: rg,
    includeHidden: X.boolean(),
    authorizedSheetIds: X.array(Ws).max(1e3),
    sourceCells: X.array(
      X.object({
        sheetId: Ws,
        address: X.string().regex(/^[A-Z]{1,3}[1-9]\d{0,6}$/),
      }).strict(),
    ).max(4e3),
    rows: X.array(
      X.object({
        row: X.number().int().min(1).max(1048576),
        cells: X.array(jW).max(20),
      }).strict(),
    ).max(500),
    columns: X.array(X.number().int().min(0).max(16383)).max(20),
    widths: X.array(X.number().finite().min(40).max(600)).max(20),
    excludedRows: X.array(X.number().int().min(1).max(1048576)).max(500),
    object: X.unknown().optional(),
    warnings: X.array(Si(2e3)).max(50),
    hash: rg,
  }).strict(),
  VW = X.object({
    id: Ws,
    fileName: Si(255),
    fileBytes: X.number()
      .int()
      .min(1)
      .max(10 * 1024 * 1024),
    sourceSha256: rg,
    inspectedAt: IR,
    execution: X.literal("never"),
    items: X.array(
      X.object({
        path: Si(4096),
        kind: Si(100),
        bytes: X.number()
          .int()
          .nonnegative()
          .max(50 * 1024 * 1024),
        sha256: rg.optional(),
        execution: X.literal("never"),
        note: Si(2e3),
      }).strict(),
    )
      .min(1)
      .max(1e3),
    warnings: X.array(Si(2e3)).max(100),
  }).strict();
function qW(e) {
  let t = NR.parse(e);
  if ((t.kind === "object") != !!t.objectId)
    throw Error("Select an existing drawing for an object report.");
  if (t.filter && RR(t.filter.range) !== t.range)
    throw Error("The report range must match its filter range.");
  return t;
}
var AR = (e) => {
  let { hash, ...n } = e;
  return Jr(n);
};
function XW(e) {
  if (new TextEncoder().encode(JSON.stringify(e)).length > 4e6)
    throw Error("A report snapshot exceeds 4 MB.");
  let t = WW.parse(e),
    n = Qe(t.request.range);
  if ((qW(t.request), t.hash !== AR(t)))
    throw Error("Report snapshot integrity check failed.");
  let i = Array.from(
    { length: n.lastCol - n.firstCol + 1 },
    (a, s) => n.firstCol + s,
  );
  if (Jr(i) !== Jr(t.columns) || t.widths.length !== i.length)
    throw Error("Invalid report column layout.");
  let r = t.rows.map((a) => a.row),
    o = [...r, ...t.excludedRows].sort((a, s) => a - s);
  if (
    t.rows.length * i.length > 2e3 ||
    o.length !== n.lastRow - n.firstRow + 1 ||
    o.some((a, s) => a !== n.firstRow + s + 1) ||
    r.some((a, s) => s > 0 && a <= r[s - 1])
  )
    throw Error("Invalid report row accounting.");
  for (let a of t.rows) {
    if (a.cells.length !== i.length) throw Error("Invalid report cell layout.");
    a.cells.forEach((s, l) => {
      if (s.address !== `${Ue(i[l])}${a.row}`)
        throw Error("Invalid report cell address.");
      s.style && (s.style = As(s.style));
    });
  }
  if (t.object) {
    let a = {
      id: t.request.sheetId,
      name: t.sheetName,
      cells: {},
      rowCount: 1048576,
      colCount: 16384,
    };
    (Fd([t.object], a), (t.object = a.objects[0]));
  }
  if (
    (t.request.kind === "object") != !!t.object ||
    (t.object && t.object.id !== t.request.objectId)
  )
    throw Error("Invalid report drawing identity.");
  if (t.hash !== AR(t)) throw Error("Report metadata is not canonical.");
  return t;
}
function OR(e, t) {
  if (e.reportOutputs !== undefined) {
    let n = X.array(
      X.object({
        id: Ws,
        version: X.number().int().min(1).max(1e6),
        snapshot: X.unknown(),
      }).strict(),
    )
      .max(20)
      .parse(e.reportOutputs);
    if (new Set(n.map((i) => i.id)).size !== n.length)
      throw Error("Duplicate report output IDs.");
    if (new TextEncoder().encode(JSON.stringify(n)).length > zW.metadataBytes)
      throw Error("Saved reports exceed 16 MB.");
    t.reportOutputs = n.map((i) => ({ ...i, snapshot: XW(i.snapshot) }));
  }
  if (e.embeddedObjectInventories !== undefined) {
    let n = X.array(VW).max(20).parse(e.embeddedObjectInventories);
    if (
      new Set(n.map((i) => i.id)).size !== n.length ||
      JSON.stringify(n).length > 2e6
    )
      throw Error("Invalid embedded-object inventory count or size.");
    t.embeddedObjectInventories = n;
  }
}
var ng = Object.freeze({
    expressionCharacters: 4096,
    syntaxNodes: 512,
    depth: 48,
    calculatedColumns: 16,
  }),
  are = Object.freeze({
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
var HW = "ledgerglass-columnar-v1",
  Ro = Object.freeze({
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
  kc = X.string()
    .min(1)
    .max(100)
    .regex(/^[A-Za-z0-9_.-]+$/),
  Va = X.string()
    .trim()
    .min(1)
    .max(128)
    .refine((e) => !/[\u0000-\u001f\u007f]/.test(e)),
  GW = X.strictObject({
    name: Va,
    type: X.enum(["text", "number", "boolean", "date"]),
    nullable: X.boolean(),
    expression: X.string()
      .trim()
      .min(1)
      .max(ng.expressionCharacters)
      .optional(),
  }),
  ZW = X.strictObject({
    id: kc,
    name: Va,
    sheetId: X.string().min(1).max(100),
    range: X.string().min(1).max(50),
    tableId: X.string().min(1).max(100).optional(),
    columns: X.array(GW).min(1).max(Ro.columns),
    dateColumn: Va.optional(),
  }),
  ig = X.strictObject({ tableId: kc, column: Va }),
  KW = X.strictObject({
    id: kc,
    one: ig,
    many: ig,
    cardinality: X.enum(["one-to-many", "one-to-one"]),
    direction: X.literal("single"),
  }),
  JW = X.strictObject({
    id: kc,
    name: Va,
    tableId: kc,
    column: Va.optional(),
    operation: X.enum([
      "sum",
      "count",
      "countRows",
      "average",
      "min",
      "max",
      "distinctCount",
      "dax",
    ]),
    expression: X.string()
      .trim()
      .min(1)
      .max(ng.expressionCharacters)
      .optional(),
  }),
  $R = X.strictObject({
    engine: X.literal(HW),
    name: Va,
    tables: X.array(ZW).min(1).max(Ro.tables),
    relationships: X.array(KW).max(Ro.relationships),
    measures: X.array(JW).max(Ro.measures),
  }),
  YW = X.union([
    X.string().max(32767),
    X.number().finite(),
    X.boolean(),
    X.null(),
  ]),
  QW = X.strictObject({
    ...ig.shape,
    operator: X.enum([
      "equals",
      "notEquals",
      "gt",
      "gte",
      "lt",
      "lte",
      "blank",
      "notBlank",
    ]),
    value: YW.optional(),
  }),
  PR = X.strictObject({
    groupBy: X.array(ig).max(Ro.groupFields),
    filters: X.array(QW).max(Ro.filters),
    measureIds: X.array(kc).max(Ro.measures),
  }),
  Vs = (e, t) => {
    if (new Set(e.map((n) => n.toLowerCase())).size !== e.length)
      throw Error(`${t} must be unique, including case variants.`);
  };
function MR(e) {
  let t = $R.safeParse(e);
  if (!t.success)
    throw Error(
      `Invalid data model: ${t.error.issues[0]?.path.join(".")}: ${t.error.issues[0]?.message}`,
    );
  let n = t.data;
  if (
    (Vs(
      n.tables.map((r) => r.id),
      "Model table IDs",
    ),
    Vs(
      n.tables.map((r) => r.name),
      "Model table names",
    ),
    Vs(
      n.relationships.map((r) => r.id),
      "Relationship IDs",
    ),
    Vs(
      n.measures.map((r) => r.id),
      "Measure IDs",
    ),
    Vs(
      n.measures.map((r) => r.name),
      "Measure names",
    ),
    n.tables.reduce(
      (r, o) => r + o.columns.filter((a) => a.expression !== undefined).length,
      0,
    ) > ng.calculatedColumns)
  )
    throw Error("The model supports at most 16 DAX calculated columns.");
  for (let r of n.tables) {
    if (
      (Vs(
        r.columns.map((o) => o.name),
        `Columns in ${r.name}`,
      ),
      !r.columns.some((o) => o.expression === undefined))
    )
      throw Error(
        "A model table requires at least one physical source column.",
      );
    if (
      r.dateColumn &&
      !r.columns.some(
        (o) =>
          o.name === r.dateColumn &&
          o.type === "date" &&
          !o.nullable &&
          o.expression === undefined,
      )
    )
      throw Error(
        "A marked calendar requires a nonnullable physical date column.",
      );
  }
  let i = (r) => {
    let o = n.tables.find((s) => s.id === r.tableId),
      a = o?.columns.find((s) => s.name === r.column);
    if (!a) throw Error(`Unknown model field ${r.tableId}.${r.column}.`);
    return a;
  };
  for (let r of n.relationships) {
    if (r.one.tableId === r.many.tableId)
      throw Error(
        "Self relationships are unsupported; use distinct role-playing tables.",
      );
    if (i(r.one).type !== i(r.many).type)
      throw Error(`Relationship ${r.id} requires matching key types.`);
    if (i(r.one).expression !== undefined || i(r.many).expression !== undefined)
      throw Error(
        "Relationship keys must be physical source columns, not DAX calculated columns.",
      );
  }
  for (let r of n.measures) {
    if (!n.tables.some((o) => o.id === r.tableId))
      throw Error(`Unknown measure table ${r.tableId}.`);
    if (r.operation === "dax") {
      if (!r.expression || r.column !== undefined)
        throw Error(
          "A DAX measure requires an expression and no aggregate column.",
        );
    } else {
      if (r.expression !== undefined)
        throw Error(
          "Unrecognized expression for an aggregate measure; only DAX measures may contain it.",
        );
      if (r.operation === "countRows") {
        if (r.column !== undefined)
          throw Error("Count rows does not take a column.");
      } else {
        let o = i({ tableId: r.tableId, column: r.column ?? "" });
        if (
          ["sum", "average", "min", "max"].includes(r.operation) &&
          o.type !== "number"
        )
          throw Error(`Measure ${r.name} requires a number column.`);
      }
    }
  }
  return n;
}
function eV(e, t) {
  let n = PR.safeParse(e);
  if (!n.success)
    throw Error(`Invalid model query: ${n.error.issues[0]?.message}`);
  let i = n.data;
  (Vs(i.measureIds, "Selected measures"),
    Vs(
      i.groupBy.map((r) => JSON.stringify(r)),
      "Group fields",
    ));
  for (let r of [...i.groupBy, ...i.filters])
    if (
      !t.tables
        .find((o) => o.id === r.tableId)
        ?.columns.some((o) => o.name === r.column)
    )
      throw Error(`Unknown query field ${r.tableId}.${r.column}.`);
  if (i.measureIds.some((r) => !t.measures.some((o) => o.id === r)))
    throw Error("A selected model measure is missing.");
  return i;
}
var ak = X.string().regex(/^[a-f0-9]{64}$/),
  tV = X.strictObject({
    definition: $R,
    query: PR,
    sourceHash: ak,
    definitionHash: ak,
    outputHash: ak,
    outputRange: X.string().min(1).max(50),
    createdAt: X.string().datetime(),
    sourceRevision: X.number().int().nonnegative(),
    includeHidden: X.boolean(),
    totals: X.array(X.number().finite().nullable()).max(Ro.measures),
  });
function DR(e, t) {
  let n = e,
    i = 0;
  if (n.dataModel !== undefined) {
    if (((i += JSON.stringify(n.dataModel).length), i > Ro.metadataCharacters))
      throw Error("Data model metadata exceeds 2,000,000 characters.");
    t.dataModel = MR(n.dataModel);
  }
  for (let r of t.sheets) {
    let o = n.sheets.find((l) => l.id === r.id)?.dataModelReport;
    if (o === undefined) continue;
    if (((i += JSON.stringify(o).length), i > Ro.metadataCharacters))
      throw Error("Data model metadata exceeds 2,000,000 characters.");
    let a = tV.safeParse(o);
    if (!a.success)
      throw Error(`Invalid model report: ${a.error.issues[0]?.message}`);
    let s = a.data;
    if (
      ((s.definition = MR(s.definition)),
      (s.query = eV(s.query, s.definition)),
      s.definitionHash !== Jr(s.definition) ||
        s.totals.length !== s.query.measureIds.length)
    )
      throw Error(
        "Model report definition or measure metadata is inconsistent.",
      );
    r.dataModelReport = s;
  }
}
var an = Object.freeze({
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
  LR = ["json", "xml", "pdf", "png", "jpeg", "html"],
  yn = (e) => X.string().max(e),
  UR = X.string().regex(/^[a-f0-9]{64}$/),
  FR = X.union([yn(32767), X.number().finite(), X.boolean(), X.null()]),
  rV = X.object({
    page: X.number().int().min(1).max(1e4),
    x: X.number().min(0).max(1),
    y: X.number().min(0).max(1),
    width: X.number().positive().max(1),
    height: X.number().positive().max(1),
  })
    .strict()
    .refine(
      (e) => e.x + e.width <= 1.000001 && e.y + e.height <= 1.000001,
      "Source region must fit its page.",
    );
function nV(e, t) {
  if (typeof e != "string") return false;
  try {
    let n = JSON.parse(e);
    return t === "array"
      ? Array.isArray(n)
      : !!n && typeof n == "object" && !Array.isArray(n);
  } catch {
    return false;
  }
}
var iV = X.object({
    value: FR,
    type: X.enum(["string", "number", "boolean", "null", "array", "object"]),
    path: yn(2048),
    lexical: yn(32767).optional(),
    declaredType: yn(200).optional(),
    missing: X.boolean().optional(),
    confidence: X.number().min(0).max(1),
    region: rV.optional(),
  })
    .strict()
    .refine(
      (e) =>
        e.type === "array" || e.type === "object"
          ? nV(e.value, e.type)
          : e.type === "null"
            ? e.value === null
            : typeof e.value === e.type,
      "Native value and type disagree.",
    )
    .refine(
      (e) => !e.missing || (e.value === null && e.type === "null"),
      "Missing fields must be null placeholders.",
    ),
  zR = X.object({
    id: yn(200).min(1),
    name: yn(200).min(1),
    path: yn(2048),
    ambiguous: X.boolean(),
    columns: X.array(
      X.object({ name: yn(200).min(1), path: yn(2048) }).strict(),
    )
      .min(1)
      .max(an.columns),
    rows: X.array(X.array(iV).max(an.columns)).min(1).max(an.rows),
  })
    .strict()
    .superRefine((e, t) => {
      (new Set(e.columns.map((n) => n.path)).size !== e.columns.length ||
        e.rows.some((n) => n.length !== e.columns.length)) &&
        t.addIssue({
          code: "custom",
          message: "Table columns and rows must align with unique paths.",
        });
    }),
  og = X.object({
    fileName: yn(255)
      .min(1)
      .refine((e) => !/[\\/\x00-\x1f\x7f\u202a-\u202e\u2066-\u2069]/.test(e)),
    format: X.enum(LR),
    byteLength: X.number().int().min(1).max(an.bytes),
    sha256: UR,
    base64: yn(Math.ceil(an.bytes / 3) * 4),
  }).strict(),
  BR = X.object({
    schemaVersion: X.literal(1),
    fileName: og.shape.fileName,
    format: X.enum(LR),
    byteLength: og.shape.byteLength,
    sha256: UR,
    tables: X.array(zR).max(an.tables),
    method: X.enum(["native-json", "native-xml", "document-adapter"]),
    extractor: yn(200).min(1),
    pageCount: X.number().int().min(1).max(1e4).optional(),
    warnings: X.array(yn(2048)).max(30),
  }).strict(),
  jR = X.object({
    decision: X.enum(["accept", "exclude", "correct"]),
    value: FR.optional(),
  })
    .strict()
    .refine(
      (e) => e.decision !== "correct" || e.value !== undefined,
      "A corrected value is required.",
    ),
  oV = X.object({
    id: yn(100).min(1),
    sheetId: yn(100).min(1),
    sheetName: yn(31).min(1),
    importedAt: X.string().datetime(),
    original: og,
    method: BR.shape.method,
    extractor: yn(200).min(1),
    pageCount: X.number().int().min(1).max(1e4).optional(),
    table: zR,
    resolutions: X.record(X.string().regex(/^\d+:\d+$/), jR),
    ambiguityReviewed: X.boolean(),
  }).strict();
function zd(e) {
  let t = BR.parse(e),
    n = 0;
  if (new Set(t.tables.map((i) => i.id)).size !== t.tables.length)
    throw Error("Source table IDs must be unique.");
  for (let i of t.tables)
    for (let r of i.rows)
      for (let o of r) {
        if (++n > an.cells) throw Error("Source cell budget exceeded.");
        if (
          t.method === "document-adapter" &&
          (!o.region ||
            !t.pageCount ||
            o.region.page > t.pageCount ||
            (["png", "jpeg"].includes(t.format) && t.pageCount !== 1))
        )
          throw Error(
            "Document cells require a source page and bounding region.",
          );
      }
  if (
    (t.format === "json" && t.method !== "native-json") ||
    (t.format === "xml" && t.method !== "native-xml") ||
    (!["json", "xml"].includes(t.format) && t.method !== "document-adapter")
  )
    throw Error("Source format and extraction method disagree.");
  return t;
}
function ck(e) {
  let t = og.parse(e),
    n = t.fileName.split(".").at(-1)?.toLowerCase();
  if ((n === "jpg" ? "jpeg" : n === "htm" ? "html" : n) !== t.format)
    throw Error("Source extension and format disagree.");
  if (
    t.base64.length !== Math.ceil(t.byteLength / 3) * 4 ||
    !/^[A-Za-z0-9+/]*={0,2}$/.test(t.base64)
  )
    throw Error("Invalid retained source encoding.");
  let i = atob(t.base64);
  if (btoa(i) !== t.base64 || i.length !== t.byteLength)
    throw Error("Retained source byte length changed.");
  let r = Uint8Array.from(i, (o) => o.charCodeAt(0));
  if (rr(r) !== t.sha256) throw Error("Retained source checksum mismatch.");
  return r;
}
function sV(e, t) {
  if (e.confidence < an.confidence && !t)
    throw Error("Review every low-confidence cell before importing.");
  if (!t) return e.value;
  let n = jR.parse(t);
  return n.decision === "exclude"
    ? null
    : n.decision === "correct"
      ? n.value
      : e.value;
}
var lk = new WeakSet();
function WR(e, t) {
  if (e.sourceImports === undefined) return;
  if (!Array.isArray(e.sourceImports) || e.sourceImports.length > an.imports)
    throw Error("Source import count exceeded.");
  if (
    e.sourceImports.some((r) => !r || typeof r != "object" || !lk.has(r)) &&
    new TextEncoder().encode(JSON.stringify(e.sourceImports)).length >
      24 * 1024 * 1024
  )
    throw Error("Source provenance metadata budget exceeded.");
  let n = 0,
    i = new Set();
  if (
    ((t.sourceImports = e.sourceImports.map((r) => {
      if (r && typeof r == "object" && lk.has(r)) {
        let s = r;
        if (((n += s.original.byteLength), i.has(s.id)))
          throw Error("Duplicate source import ID.");
        return (i.add(s.id), s);
      }
      let o = oV.parse(r);
      if (
        (ck(o.original),
        zd({
          schemaVersion: 1,
          fileName: o.original.fileName,
          format: o.original.format,
          byteLength: o.original.byteLength,
          sha256: o.original.sha256,
          tables: [o.table],
          method: o.method,
          extractor: o.extractor,
          ...(o.pageCount ? { pageCount: o.pageCount } : {}),
          warnings: [],
        }),
        o.table.ambiguous && !o.ambiguityReviewed)
      )
        throw Error("Ambiguous source tables require review.");
      let a = new Set();
      if (
        (o.table.rows.forEach((s, l) =>
          s.forEach((c, u) => {
            let d = `${l}:${u}`;
            (a.add(d), sV(c, o.resolutions[d]));
          }),
        ),
        Object.keys(o.resolutions).some((s) => !a.has(s)))
      )
        throw Error("Unknown source review cell.");
      if (i.has(o.id)) throw Error("Duplicate source import ID.");
      return (i.add(o.id), (n += o.original.byteLength), VR(o), lk.add(o), o);
    })),
    n > an.retainedBytes)
  )
    throw Error("Retained source byte budget exceeded.");
}
function VR(e) {
  Object.freeze(e);
  for (let t of Object.values(e)) t && typeof t == "object" && VR(t);
}
var sg = Object.freeze({
    maxBytes: 10 * 1024 * 1024,
    maxParts: 2e3,
    maxDetails: 1e5,
  }),
  uk = new WeakSet();
function Bd(e) {
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
    e.byteLength > sg.maxBytes ||
    typeof e.sha256 != "string" ||
    !/^[a-f0-9]{64}$/.test(e.sha256) ||
    typeof e.base64 != "string" ||
    e.base64.length !== Math.ceil(e.byteLength / 3) * 4 ||
    !/^[A-Za-z0-9+/]*={0,2}$/.test(e.base64)
  )
    throw Error(
      "Invalid retained XLSX source metadata or 10 MiB source limit exceeded.",
    );
  let t;
  try {
    t = atob(e.base64);
  } catch {
    throw Error("Invalid retained XLSX source encoding.");
  }
  if (btoa(t) !== e.base64 || t.length !== e.byteLength)
    throw Error("Retained XLSX source byte length or encoding changed.");
  let n = Uint8Array.from(t, (i) => i.charCodeAt(0));
  if (rr(n) !== e.sha256)
    throw Error("Retained XLSX source checksum does not match its bytes.");
  return n;
}
function ag(e, t) {
  if (e.byteLength > sg.maxBytes)
    throw Error("Original XLSX exceeds the 10 MiB retention limit.");
  let n = "";
  for (let r = 0; r < e.length; r += 8192)
    n += String.fromCharCode(...e.subarray(r, r + 8192));
  let i = {
    schemaVersion: 1,
    fileName: t,
    byteLength: e.length,
    sha256: rr(e),
    base64: btoa(n),
  };
  return (Bd(i), Object.freeze(i), uk.add(i), i);
}
function qR(e, t) {
  if (e.xlsxSource === undefined) return;
  let n = e.xlsxSource;
  if (n && typeof n == "object" && uk.has(n)) {
    t.xlsxSource = n;
    return;
  }
  (Bd(n),
    (t.xlsxSource = Object.freeze({
      schemaVersion: 1,
      fileName: n.fileName,
      byteLength: n.byteLength,
      sha256: n.sha256,
      base64: n.base64,
    })),
    uk.add(t.xlsxSource));
}
function XR(e) {
  if (!e.xlsxSource)
    throw Error(
      "This workbook has no retained XLSX source. Open an XLSX file to retain its original bytes.",
    );
  return {
    bytes: Bd(e.xlsxSource),
    fileName: e.xlsxSource.fileName,
    sha256: e.xlsxSource.sha256,
  };
}
var _c = ["xls", "xlsb", "xlsm", "xltx", "xltm", "ods"],
  HR = (e) =>
    typeof e == "string" &&
    e.length > 0 &&
    e.length <= 255 &&
    !/[\\/\x00-\x1f\x7f\u202a-\u202e\u2066-\u2069]/.test(e),
  GR = new WeakSet();
function lg(e) {
  if (
    !e ||
    typeof e != "object" ||
    !_c.includes(e.format) ||
    !HR(e.fileName) ||
    !e.fileName.toLowerCase().endsWith("." + e.format) ||
    !e.conversion ||
    !["template-adapter", "external-conversion"].includes(
      e.conversion.method,
    ) ||
    (e.conversion.method === "template-adapter" && e.format !== "xltx") ||
    !HR(e.conversion.convertedFileName) ||
    !/\.xlsx$/i.test(e.conversion.convertedFileName) ||
    typeof e.conversion.convertedSha256 != "string" ||
    !/^[a-f0-9]{64}$/.test(e.conversion.convertedSha256) ||
    typeof e.conversion.reviewedAt != "string" ||
    !/^\d{4}-\d\d-\d\dT\d\d:\d\d:\d\d\.\d{3}Z$/.test(e.conversion.reviewedAt) ||
    !Number.isFinite(Date.parse(e.conversion.reviewedAt)) ||
    e.conversion.equivalence !== "not-verified"
  )
    throw Error("Invalid alternate workbook source or conversion provenance.");
  return Bd({ ...e, fileName: "retained.xlsx" });
}
function dk(e, t, n, i) {
  let r = {
    ...ag(e, "retained.xlsx"),
    fileName: t,
    format: n,
    conversion: { ...i },
  };
  return (lg(r), Object.freeze(r.conversion), Object.freeze(r), GR.add(r), r);
}
function ZR(e, t) {
  if (e.alternateSource === undefined) return;
  let n = e.alternateSource;
  if (GR.has(n)) {
    t.alternateSource = n;
    return;
  }
  let i = lg(n);
  t.alternateSource = dk(i, n.fileName, n.format, n.conversion);
}
var jn = Object.freeze({
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
  aV = [
    "vba",
    "office-script",
    "office-script-container",
    "excel-container",
    "legacy-container",
  ];
function lV(e) {
  if (
    typeof e != "string" ||
    e.length > Math.ceil(jn.maxArtifactBytes / 3) * 4 ||
    !/^(?:[A-Za-z0-9+/]{4})*(?:[A-Za-z0-9+/]{2}==|[A-Za-z0-9+/]{3}=)?$/.test(e)
  )
    throw Error("Invalid automation artifact encoding or size.");
  let t = atob(e);
  if (btoa(t) !== e)
    throw Error("Automation artifact is not canonical base64.");
  return Uint8Array.from(t, (n) => n.charCodeAt(0));
}
var as = (e, t, n) => {
    if (typeof e != "string" || !e.trim() || e.length > t)
      throw Error(`Invalid automation ${n}.`);
    return e;
  },
  fk = (e) => {
    if (typeof e != "string" || !/^[a-f0-9]{64}$/.test(e))
      throw Error("Invalid automation source hash.");
    return e;
  },
  jd = (e, t, n) => {
    if (!Number.isSafeInteger(e) || Number(e) < t || Number(e) > n)
      throw Error("Invalid automation evidence count.");
    return Number(e);
  },
  KR = (e) => {
    let t = as(e, 40, "timestamp");
    if (!/^\d{4}-\d\d-\d\dT/.test(t) || !Number.isFinite(Date.parse(t)))
      throw Error("Invalid automation timestamp.");
    return t;
  },
  cV = (e) => {
    let t = as(e, 40, "approved range"),
      n = /^([A-Z]{1,3})([1-9]\d*):([A-Z]{1,3})([1-9]\d*)$/.exec(t);
    if (!n) throw Error("Invalid automation approved range.");
    let i = (l) => [...l].reduce((c, u) => c * 26 + u.charCodeAt(0) - 64, 0),
      [r, o, a, s] = [i(n[1]), Number(n[2]), i(n[3]), Number(n[4])];
    if (
      r > a ||
      o > s ||
      a > 16384 ||
      s > 1048576 ||
      (a - r + 1) * (s - o + 1) > jn.maxCells
    )
      throw Error("Invalid automation approved range.");
    return t;
  };
function cg(e) {
  let t = as(e, 255, "file name");
  if (/[\\/\x00-\x1f\x7f]/.test(t) || t === "." || t === "..")
    throw Error("Invalid automation file name.");
  return t;
}
function JR(e, t) {
  if (e.automationArtifacts !== undefined) {
    if (
      !Array.isArray(e.automationArtifacts) ||
      e.automationArtifacts.length > jn.maxArtifacts
    )
      throw Error("Keep at most 20 automation artifacts.");
    let n = 0,
      i = new Set();
    t.automationArtifacts = e.automationArtifacts.map((r) => {
      if (!r || typeof r != "object" || Array.isArray(r))
        throw Error("Invalid automation artifact.");
      let o = r,
        a = as(o.id, 100, "artifact ID");
      if (i.has(a)) throw Error("Duplicate automation artifact ID.");
      i.add(a);
      let s = lV(o.base64);
      if (
        ((n += s.length),
        !s.length ||
          s.length !== o.byteLength ||
          s.length > jn.maxArtifactBytes ||
          n > jn.maxStoredBytes)
      )
        throw Error(
          "Automation artifacts exceed their 2 MB per-file or 10 MB retained-byte budget.",
        );
      if (rr(s) !== fk(o.sha256))
        throw Error(
          "Automation artifact checksum does not match its preserved bytes.",
        );
      if (!aV.includes(o.kind))
        throw Error("Unknown automation artifact kind.");
      return {
        id: a,
        fileName: cg(o.fileName),
        kind: o.kind,
        sha256: o.sha256,
        byteLength: s.length,
        base64: o.base64,
        retainedAt: KR(o.retainedAt),
        retainedBy: as(o.retainedBy, 100, "retainer"),
      };
    });
  }
  if (e.automationRuns !== undefined) {
    if (
      !Array.isArray(e.automationRuns) ||
      e.automationRuns.length > jn.maxRuns
    )
      throw Error("Automation run evidence is limited to 100 runs.");
    let n = new Set();
    t.automationRuns = e.automationRuns.map((i) => {
      if (!i || typeof i != "object" || Array.isArray(i))
        throw Error("Invalid automation run evidence.");
      let r = i,
        o = as(r.id, 100, "run ID");
      if (n.has(o)) throw Error("Duplicate automation run ID.");
      if (
        (n.add(o),
        !["vba", "office-script"].includes(r.language) ||
          r.adapter !== "literal-ranges-v1")
      )
        throw Error("Unknown automation migration adapter.");
      if (
        !Array.isArray(r.sourceLines) ||
        r.sourceLines.length > jn.maxCommands
      )
        throw Error("Invalid automation source lines.");
      let a = jd(r.beforeRevision, 0, Number.MAX_SAFE_INTEGER - 2),
        s = jd(r.afterRevision, 1, Number.MAX_SAFE_INTEGER - 1);
      if (s !== a + 1)
        throw Error("Automation run revisions are inconsistent.");
      if (r.sourceLines.length !== r.commands)
        throw Error(
          "Automation source-line evidence must match the command count.",
        );
      return {
        id: o,
        eventId: as(r.eventId, 100, "event ID"),
        sourceName: cg(r.sourceName),
        sourceSha256: fk(r.sourceSha256),
        codeSha256: fk(r.codeSha256),
        language: r.language,
        adapter: r.adapter,
        actor: as(r.actor, 100, "actor"),
        timestamp: KR(r.timestamp),
        beforeRevision: a,
        afterRevision: s,
        sheetId: as(r.sheetId, 100, "sheet ID"),
        range: cV(r.range),
        commands: jd(r.commands, 1, jn.maxCommands),
        changedCells: jd(r.changedCells, 0, jn.maxCells),
        sourceLines: r.sourceLines.map((l) => jd(l, 1, jn.maxSourceBytes + 1)),
      };
    });
  }
}
function YR(e, t) {
  if (e === undefined) return;
  if (!e || typeof e != "object" || Array.isArray(e))
    throw Error("Invalid delimited source footprint.");
  let n = e;
  if (
    !Number.isInteger(n.records) ||
    Number(n.records) < 1 ||
    Number(n.records) > Ne.maxRows ||
    !Number.isInteger(n.columns) ||
    Number(n.columns) < 1 ||
    Number(n.columns) > Ne.maxCols ||
    typeof n.hasHeader != "boolean"
  )
    throw Error(
      "Delimited source footprint exceeds worksheet bounds or has invalid header metadata.",
    );
  t.delimitedSource = {
    records: Number(n.records),
    columns: Number(n.columns),
    hasHeader: n.hasHeader,
  };
}
var uV = 100,
  dV = [
    "deleted-cells",
    "deleted-sheet",
    "overwritten",
    "missing-sheet",
    "out-of-bounds",
  ],
  QR = (e) =>
    !!e &&
    typeof e == "object" &&
    !Array.isArray(e) &&
    [Object.prototype, null].includes(Object.getPrototypeOf(e)),
  eN = (e, t) => Object.keys(e).every((n) => t.includes(n)),
  hk = (e, t) =>
    typeof e == "string" &&
    !!e.trim() &&
    e.length <= t &&
    !/[\u0000-\u001f]/.test(e);
function fV(e) {
  if (typeof e != "string" || !/^\$?[A-Za-z]{1,3}\$?[1-9]\d{0,6}$/.test(e))
    throw Error("A formula watch must identify one valid Excel cell address.");
  let t = tt(e.replace(/\$/g, "").toUpperCase());
  if (t.row >= Ne.maxRows || t.col >= Ne.maxCols)
    throw Error("A formula watch address is outside Excel's cell bounds.");
  return `${Ue(t.col)}${t.row + 1}`;
}
function tN(e, t) {
  if (e === undefined) return [];
  if (!Array.isArray(e) || e.length > uV)
    throw Error("A workbook may contain at most 100 formula watches.");
  let n = new Set(),
    i = new Set();
  return e.map((r) => {
    if (
      !QR(r) ||
      !eN(r, ["id", "anchor", "orphaned"]) ||
      !hk(r.id, 100) ||
      !QR(r.anchor) ||
      !eN(r.anchor, ["sheetId", "sheetName", "address"]) ||
      !hk(r.anchor.sheetId, 100) ||
      !hk(r.anchor.sheetName, 31)
    )
      throw Error("Invalid formula watch metadata.");
    if (n.has(r.id)) throw Error("Formula watch IDs must be unique.");
    if ((n.add(r.id), r.orphaned !== undefined && !dV.includes(r.orphaned)))
      throw Error("Invalid formula watch orphan reason.");
    let o = r.anchor,
      a = fV(o.address),
      s = t.sheets.find((u) => u.id === o.sheetId),
      l = r.orphaned ?? (s ? undefined : "missing-sheet"),
      c = JSON.stringify([r.anchor.sheetId, a]);
    if (!l && i.has(c))
      throw Error("Active formula watches must not repeat the same cell.");
    return (
      l || i.add(c),
      {
        id: r.id,
        anchor: {
          sheetId: r.anchor.sheetId,
          sheetName: !l && s ? s.name : r.anchor.sheetName,
          address: a,
        },
        ...(l ? { orphaned: l } : {}),
      }
    );
  });
}
var Tn = {
  inputs: 32,
  outputs: 16,
  scenarios: 100,
  axisValues: 100,
  trials: 400,
  dependencyCells: 500,
  workUnits: 5e4,
  savedAnalyses: 40,
};
var qs = (e) => JSON.stringify([e.sheetId, e.address]),
  nN = (e) => typeof e == "number" && Number.isFinite(e);
function ls(e, t) {
  if (
    !e ||
    typeof e != "object" ||
    Array.isArray(e) ||
    ![Object.prototype, null].includes(Object.getPrototypeOf(e)) ||
    Object.keys(e).some((n) => !t.includes(n))
  )
    throw Error("Invalid what-if analysis fields.");
  return e;
}
function Xs(e, t = 100) {
  if (
    typeof e != "string" ||
    !e.trim() ||
    e.length > t ||
    /[\u0000-\u001f]/.test(e)
  )
    throw Error(
      `What-if names must be nonempty text of at most ${t} characters.`,
    );
  return e.trim();
}
function Sc(e, t, n = 1) {
  if (!Array.isArray(e) || e.length < n || e.length > t)
    throw Error(`What-if lists require ${n}\u2013${t} entries.`);
  return Array.from(e);
}
function Ec(e, t) {
  return Sc(e, t).map((n) => {
    if (!nN(n))
      throw Error("Every what-if input value must be a finite number.");
    return n;
  });
}
function iN(e) {
  let t = ls(e, ["sheetId", "address"]),
    n = Xs(t.sheetId);
  if (typeof t.address != "string")
    throw Error("A what-if cell needs an Excel address.");
  let i = tt(t.address);
  return { sheetId: n, address: `${Ue(i.col)}${i.row + 1}` };
}
function rN(e, t) {
  let n = Sc(e, t).map(iN);
  if (new Set(n.map(qs)).size !== n.length)
    throw Error("What-if cell references must not repeat.");
  return n;
}
function hV(e) {
  let t = ls(e, [
      "name",
      "kind",
      "inputs",
      "outputs",
      "scenarios",
      "orientation",
      "values",
      "rowValues",
      "columnValues",
    ]),
    n = {
      name: Xs(t.name),
      inputs: rN(t.inputs, Tn.inputs),
      outputs: rN(t.outputs, Tn.outputs),
    };
  if (n.inputs.some((r) => n.outputs.some((o) => qs(r) === qs(o))))
    throw Error("Changing inputs and formula results must be different cells.");
  let i = ["name", "kind", "inputs", "outputs"];
  if (t.kind === "scenarios") {
    ls(t, [...i, "scenarios"]);
    let r = Sc(t.scenarios, Tn.scenarios).map((o) => {
      let a = ls(o, ["name", "values"]),
        s = Ec(a.values, Tn.inputs);
      if (s.length !== n.inputs.length)
        throw Error(
          "Each scenario needs exactly one value per changing input.",
        );
      return { name: Xs(a.name), values: s };
    });
    if (new Set(r.map((o) => o.name.toLowerCase())).size !== r.length)
      throw Error("Scenario names must be unique, ignoring case.");
    return { ...n, kind: t.kind, scenarios: r };
  }
  if (t.kind === "one-variable") {
    if (
      (ls(t, [...i, "orientation", "values"]),
      n.inputs.length !== 1 || !["row", "column"].includes(t.orientation))
    )
      throw Error(
        "A one-variable table needs one input and a row or column orientation.",
      );
    return {
      ...n,
      kind: t.kind,
      orientation: t.orientation,
      values: Ec(t.values, Tn.axisValues),
    };
  }
  if (t.kind === "two-variable") {
    if (
      (ls(t, [...i, "rowValues", "columnValues"]),
      n.inputs.length !== 2 || n.outputs.length !== 1)
    )
      throw Error(
        "A two-variable table needs two different inputs and exactly one formula result.",
      );
    let r = Ec(t.rowValues, Tn.axisValues),
      o = Ec(t.columnValues, Tn.axisValues);
    if (r.length * o.length > Tn.trials)
      throw Error("A sensitivity table supports at most 400 combinations.");
    return { ...n, kind: t.kind, rowValues: r, columnValues: o };
  }
  throw Error(
    "Choose scenarios, a one-variable table or a two-variable table.",
  );
}
function pV(e) {
  return e.kind === "scenarios"
    ? e.scenarios.map((t) => ({ name: t.name, inputs: t.values }))
    : e.kind === "one-variable"
      ? e.values.map((t) => ({ name: String(t), inputs: [t] }))
      : e.columnValues.flatMap((t) =>
          e.rowValues.map((n) => ({ name: `${n} / ${t}`, inputs: [n, t] })),
        );
}
function oN(e) {
  let t = new Set();
  return Sc(e, Tn.savedAnalyses, 0).map((n) => {
    let i = ls(n, [
        "id",
        "reportSheetId",
        "request",
        "baseline",
        "trials",
        "assumptions",
        "sourceRevision",
        "dependencyCells",
        "workUnits",
      ]),
      r = Xs(i.id),
      o = Xs(i.reportSheetId),
      a = hV(i.request);
    if (t.has(r)) throw Error("Saved what-if IDs must be unique.");
    t.add(r);
    let s = (p) => {
        let h = ls(p, ["name", "inputs", "outputs", "error"]),
          m = Xs(h.name, 200),
          g = Ec(h.inputs, Tn.inputs),
          b = h.outputs === null ? null : Ec(h.outputs, Tn.outputs),
          w = h.error === null ? null : Xs(h.error, 2e3);
        if (
          g.length !== a.inputs.length ||
          (b && b.length !== a.outputs.length) ||
          !!b == !!w
        )
          throw Error("Invalid saved what-if result dimensions or outcome.");
        return { name: m, inputs: g, outputs: b, error: w };
      },
      l = s(i.baseline),
      c = Sc(i.trials, Tn.trials).map(s),
      u = pV(a);
    if (
      c.length !== u.length ||
      c.some(
        (p, h) =>
          p.name !== u[h].name ||
          JSON.stringify(p.inputs) !== JSON.stringify(u[h].inputs),
      )
    )
      throw Error("Saved what-if results must match their named assumptions.");
    if (l.name !== "Baseline")
      throw Error("Saved what-if baseline is missing.");
    let d = new Set(),
      f = Sc(i.assumptions, Tn.dependencyCells).map((p) => {
        let h = ls(p, ["sheetId", "sheetName", "address", "raw", "literal"]),
          m = iN({ sheetId: h.sheetId, address: h.address });
        if (d.has(qs(m)))
          throw Error("Saved assumptions must not repeat a source cell.");
        if (
          (d.add(qs(m)),
          !(
            h.raw === null ||
            typeof h.raw == "boolean" ||
            nN(h.raw) ||
            (typeof h.raw == "string" && h.raw.length <= 32767)
          ) ||
            (h.literal !== undefined && typeof h.literal != "boolean"))
        )
          throw Error("Invalid saved what-if assumption.");
        return {
          ...m,
          sheetName: Xs(h.sheetName, 31),
          raw: h.raw,
          ...(h.literal === undefined ? {} : { literal: h.literal }),
        };
      });
    if (
      [...a.inputs, ...a.outputs].some((p) => !d.has(qs(p))) ||
      a.inputs.some(
        (p, h) => f.find((m) => qs(m) === qs(p))?.raw !== l.inputs[h],
      )
    )
      throw Error(
        "Saved baseline inputs and dependency assumptions do not reconcile.",
      );
    for (let [p, h] of [
      ["sourceRevision", Number.MAX_SAFE_INTEGER - 1],
      ["dependencyCells", Tn.dependencyCells],
      ["workUnits", Tn.workUnits],
    ])
      if (!Number.isSafeInteger(i[p]) || i[p] < 0 || i[p] > h)
        throw Error("Invalid what-if evidence counts.");
    if (
      i.dependencyCells !== f.length ||
      i.workUnits !== f.length * (c.length + 1)
    )
      throw Error("What-if work counts do not reconcile.");
    return {
      id: r,
      reportSheetId: o,
      request: a,
      baseline: l,
      trials: c,
      assumptions: f,
      sourceRevision: i.sourceRevision,
      dependencyCells: f.length,
      workUnits: i.workUnits,
    };
  });
}
var pk = {
  areaCells: 5e3,
  pages: 100,
  titleRows: 100,
  titleColumns: 50,
  breaksPerAxis: 100,
  renderedCells: 1e5,
};
function sN(e, t, n) {
  if (e === undefined || e === "") return;
  if (typeof e != "string")
    throw Error("Print titles must be a single row or column interval.");
  let i = (
    t === "row"
      ? /^\$?([1-9]\d*):\$?([1-9]\d*)$/
      : /^\$?([A-Za-z]{1,3}):\$?([A-Za-z]{1,3})$/
  ).exec(e.trim());
  if (!i)
    throw Error(
      t === "row"
        ? "Repeat rows must use a range such as 1:2."
        : "Repeat columns must use a range such as A:B.",
    );
  let r = t === "row" ? Number(i[1]) - 1 : tt(i[1] + "1").col,
    o = t === "row" ? Number(i[2]) - 1 : tt(i[2] + "1").col,
    a = t === "row" ? n.rowCount : n.colCount,
    s = t === "row" ? pk.titleRows : pk.titleColumns;
  if (
    !Number.isSafeInteger(r) ||
    !Number.isSafeInteger(o) ||
    r < 0 ||
    o < r ||
    o >= a ||
    o - r + 1 > s
  )
    throw Error(
      `Print titles support at most ${s} ${t === "row" ? "rows" : "columns"} within the worksheet.`,
    );
  return { start: r, end: o };
}
function aN(e, t) {
  return t === "row"
    ? `${e.start + 1}:${e.end + 1}`
    : `${Ue(e.start)}:${Ue(e.end)}`;
}
function ug(e, t) {
  if (e !== undefined) {
    if (
      !Array.isArray(e) ||
      e.length > pk.breaksPerAxis ||
      e.some(
        (n, i) =>
          !Number.isSafeInteger(n) ||
          n < 1 ||
          n >= t ||
          (i > 0 && n <= e[i - 1]),
      )
    )
      throw Error(
        "Manual page breaks require at most 100 distinct ascending boundaries within each worksheet axis (before row/column 2 or later).",
      );
    return e.length ? [...e] : undefined;
  }
}
function bV(e) {
  return {
    orientation: "portrait",
    paper: "letter",
    margins: {
      top: 0.75,
      bottom: 0.75,
      left: 0.7,
      right: 0.7,
      header: 0.3,
      footer: 0.3,
    },
    printArea: e ? Gm(e) : "A1",
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 0,
    scale: 100,
    gridlines: true,
    headings: false,
    pageOrder: "overThenDown",
    header: "",
    footer: "",
  };
}
function Wd(e, t) {
  if (e === undefined) return;
  if (!e || typeof e != "object" || Array.isArray(e))
    throw new Error("Invalid page layout.");
  let n = e,
    i = n.margins,
    r = (b, w, _) => {
      if (typeof b != "number" || !Number.isFinite(b) || b < w || b > _)
        throw new Error("Page setup contains an out-of-range number.");
      return b;
    };
  if (
    !["portrait", "landscape"].includes(String(n.orientation)) ||
    !["letter", "a4", "legal"].includes(String(n.paper)) ||
    !i ||
    typeof i != "object"
  )
    throw new Error("Invalid paper, orientation, or margins.");
  if (
    n.pageOrder !== undefined &&
    n.pageOrder !== "overThenDown" &&
    n.pageOrder !== "downThenOver"
  )
    throw Error(
      "Unsupported print page order. Choose Over, then down or Down, then over.",
    );
  let o = Object.fromEntries(
      ["top", "bottom", "left", "right", "header", "footer"].map((b) => [
        b,
        r(i[b], 0, 3),
      ]),
    ),
    a =
      n.paper === "a4"
        ? [8.2677, 11.6929]
        : n.paper === "legal"
          ? [8.5, 14]
          : [8.5, 11],
    [s, l] = n.orientation === "landscape" ? a.reverse() : a;
  if (o.left + o.right >= s - 0.25 || o.top + o.bottom >= l - 0.25)
    throw new Error("Margins leave no printable page area.");
  if (typeof n.printArea != "string") throw new Error("Enter a print area.");
  let c = Qe(n.printArea);
  if (c.lastRow >= t.rowCount || c.lastCol >= t.colCount)
    throw new Error("Print area exceeds the worksheet.");
  if (
    ["fitToPage", "gridlines", "headings"].some((b) => typeof n[b] != "boolean")
  )
    throw new Error("Invalid print option.");
  if (
    ["header", "footer"].some(
      (b) => typeof n[b] != "string" || String(n[b]).length > 255,
    )
  )
    throw new Error("Headers and footers must be at most 255 characters.");
  let u = r(n.fitToWidth, 0, 10),
    d = r(n.fitToHeight, 0, 10),
    f = r(n.scale, 10, 400);
  if (![u, d, f].every(Number.isInteger))
    throw new Error("Page fitting and scale must be whole numbers.");
  let p = sN(n.repeatRows, "row", t),
    h = sN(n.repeatColumns, "column", t),
    m = ug(n.rowBreaks, t.rowCount),
    g = ug(n.columnBreaks, t.colCount);
  t.pageLayout = {
    orientation: n.orientation,
    paper: n.paper,
    margins: o,
    printArea: n.printArea,
    fitToPage: n.fitToPage,
    fitToWidth: u,
    fitToHeight: d,
    scale: f,
    gridlines: n.gridlines,
    headings: n.headings,
    header: n.header,
    footer: n.footer,
    ...(n.pageOrder !== undefined ? { pageOrder: n.pageOrder } : {}),
    ...(p ? { repeatRows: aN(p, "row") } : {}),
    ...(h ? { repeatColumns: aN(h, "column") } : {}),
    ...(m ? { rowBreaks: m } : {}),
    ...(g ? { columnBreaks: g } : {}),
  };
}
function lN(e, t) {
  let n = e.pageSetup,
    i = bV(t);
  if (
    n.pageOrder !== undefined &&
    n.pageOrder !== "overThenDown" &&
    n.pageOrder !== "downThenOver"
  )
    throw Error(
      "The Excel worksheet has an unsupported pageOrder value. Import stopped without changing the current workbook.",
    );
  let r = {
    ...i,
    orientation: n.orientation === "landscape" ? "landscape" : "portrait",
    paper: n.paperSize === 9 ? "a4" : n.paperSize === 5 ? "legal" : "letter",
    margins: { ...i.margins, ...n.margins },
    printArea: n.printArea?.split("&&")[0] || i.printArea,
    fitToPage: !!n.fitToPage,
    fitToWidth: Math.max(0, Math.min(10, n.fitToWidth ?? 1)),
    fitToHeight: Math.max(0, Math.min(10, n.fitToHeight ?? 0)),
    scale: Math.max(10, Math.min(400, n.scale ?? 100)),
    gridlines: !!n.showGridLines,
    headings: !!n.showRowColHeaders,
    header: String(e.headerFooter?.oddHeader ?? "").slice(0, 255),
    footer: String(e.headerFooter?.oddFooter ?? "").slice(0, 255),
    pageOrder: n.pageOrder ?? "downThenOver",
    ...(n.printTitlesRow ? { repeatRows: n.printTitlesRow } : {}),
    ...(n.printTitlesColumn ? { repeatColumns: n.printTitlesColumn } : {}),
  };
  Wd(r, t);
}
var yV = [
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
  vV = (e) =>
    !!e &&
    typeof e == "object" &&
    !Array.isArray(e) &&
    [Object.prototype, null].includes(Object.getPrototypeOf(e)),
  Hs = () => {
    throw Error(
      "Invalid or over-capacity text cleanup evidence. Keep the original native backup.",
    );
  };
function dg(e, t) {
  (!vV(e) || Object.keys(e).some((n) => !t.includes(n))) && Hs();
}
function uN(e) {
  return (
    e === null ||
    typeof e == "boolean" ||
    (typeof e == "number" && Number.isFinite(e)) ||
    (typeof e == "string" && e.length <= 32767)
  );
}
function qa(e, t) {
  return typeof e == "string" && e.length > 0 && e.length <= t;
}
function cN(e) {
  if (typeof e != "string") return false;
  try {
    let t = tt(e);
    return `${Ue(t.col)}${t.row + 1}` === e;
  } catch {
    return false;
  }
}
function wV(e) {
  if (!Array.isArray(e) || !e.length || e.length > 12)
    throw Error("Choose 1\u201312 text cleanup steps.");
  return e.map(
    (t) => (
      dg(t, ["function", "args"]),
      (!yV.includes(t.function) ||
        !Array.isArray(t.args) ||
        t.args.length > 5 ||
        t.args.some((n) => n === null || !uN(n)) ||
        JSON.stringify(t.args).length > 4096) &&
        Hs(),
      { function: t.function, args: [...t.args] }
    ),
  );
}
function dN(e, t) {
  if (e === undefined) return;
  (!Array.isArray(e) || e.length > 100 || JSON.stringify(e).length > 5e6) &&
    Hs();
  let n = new Set(),
    i = 0;
  t.textCleanups = e.map((r) => {
    dg(r, [
      "id",
      "workbookId",
      "sourceRevision",
      "sourceSheetId",
      "sourceSheetName",
      "outputSheetId",
      "outputSheetName",
      "createdAt",
      "steps",
      "semantics",
      "rows",
    ]);
    for (let c of ["id", "workbookId", "sourceSheetId", "outputSheetId"])
      qa(r[c], 100) || Hs();
    ((!qa(r.sourceSheetName, 31) ||
      !qa(r.outputSheetName, 31) ||
      n.has(r.id) ||
      !Number.isSafeInteger(r.sourceRevision) ||
      Number(r.sourceRevision) < 0 ||
      r.semantics !== "unicode-v2-en-US-explicit-separators" ||
      !Array.isArray(r.rows) ||
      !r.rows.length ||
      r.rows.length > 500 ||
      (i += r.rows.length) > 1e4 ||
      typeof r.createdAt != "string" ||
      !/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}Z$/.test(r.createdAt) ||
      !Number.isFinite(Date.parse(r.createdAt)) ||
      new Date(r.createdAt).toISOString() !== r.createdAt) &&
      Hs(),
      n.add(r.id));
    let o = wV(r.steps),
      a = new Set(),
      s = new Set(),
      l = r.rows.map((c) => {
        (dg(c, [
          "sourceAddress",
          "original",
          "result",
          "outputAddress",
          "origin",
        ]),
          (!cN(c.sourceAddress) ||
            !cN(c.outputAddress) ||
            a.has(c.sourceAddress) ||
            s.has(c.outputAddress) ||
            !uN(c.result)) &&
            Hs(),
          a.add(c.sourceAddress),
          s.add(c.outputAddress));
        let u = c.original === undefined ? undefined : Oi(c.original);
        u &&
          typeof u.raw == "string" &&
          u.raw.startsWith("=") &&
          !u.literal &&
          Hs();
        let d;
        if (c.origin !== undefined) {
          dg(c.origin, [
            "id",
            "sourceId",
            "sourceName",
            "sourceSheet",
            "originalRow",
            "kind",
          ]);
          let f = c.origin;
          ((!qa(f.id, 100) ||
            !qa(f.sourceId, 100) ||
            !qa(f.sourceName, 255) ||
            !qa(f.sourceSheet, 31) ||
            !Number.isInteger(f.originalRow) ||
            Number(f.originalRow) < 1 ||
            Number(f.originalRow) > 1048576 ||
            !["snapshot", "file"].includes(String(f.kind))) &&
            Hs(),
            (d = { ...f }));
        }
        return {
          sourceAddress: c.sourceAddress,
          outputAddress: c.outputAddress,
          result: c.result,
          ...(u ? { original: u } : {}),
          ...(d ? { origin: d } : {}),
        };
      });
    return { ...r, steps: o, rows: l };
  });
}
var fN = {
    sourceCells: 500,
    sourceCharacters: 2e5,
    fields: 50,
    steps: 12,
    outputCells: 5e3,
    outputCharacters: 1e6,
    inferenceColumns: 8,
    inferencePatterns: 8e3,
    inferenceCandidates: 30,
    records: 100,
    evidenceCharacters: 5e6,
  },
  fg = X.number().int().min(0).max(49),
  cs = X.array(fg)
    .min(1)
    .max(50)
    .refine(
      (e) => new Set(e).size === e.length,
      "Choose distinct column positions.",
    ),
  hN = X.string().max(32767),
  hg = X.union([hN, X.number().finite(), X.boolean(), X.null()]),
  pN = X.enum(["none", "upper", "lower", "proper"]),
  xV = X.strictObject({
    column: fg,
    mode: X.enum(["whole", "token", "slice", "before", "after"]),
    delimiter: X.string().min(1).max(16).default(" "),
    index: X.number().int().min(-50).max(49).default(0),
    start: X.number().int().min(0).max(32767).default(0),
    length: X.number().int().min(1).max(32767).default(1),
  }),
  kV = X.strictObject({
    parts: X.array(xV).min(1).max(8),
    separator: X.string().max(100),
    prefix: X.string().max(1e3),
    suffix: X.string().max(1e3),
    casing: pN,
  }),
  _V = X.strictObject({ row: X.number().int().min(1).max(1048576), value: hN }),
  EV = X.array(_V)
    .min(2)
    .max(10)
    .refine(
      (e) => new Set(e.map((t) => t.row)).size === e.length,
      "Use examples from distinct source rows.",
    ),
  SV = X.discriminatedUnion("kind", [
    X.strictObject({
      kind: X.literal("trim"),
      columns: cs,
      mode: X.enum(["edges", "collapse"]),
    }),
    X.strictObject({ kind: X.literal("clean"), columns: cs }),
    X.strictObject({ kind: X.literal("case"), columns: cs, mode: pN }),
    X.strictObject({
      kind: X.literal("replace"),
      columns: cs,
      find: X.string().min(1).max(1e3),
      replacement: X.string().max(1e3),
    }),
    X.strictObject({
      kind: X.literal("split"),
      column: fg,
      mode: X.enum(["delimited", "fixed"]),
      delimiters: X.string().max(16),
      collapse: X.boolean(),
      breaks: X.array(X.number().int().min(1).max(32766)).max(49),
    }),
    X.strictObject({
      kind: X.literal("concat"),
      columns: cs,
      separator: X.string().max(100),
      name: X.string().min(1).max(100),
      skipBlank: X.boolean(),
    }),
    X.strictObject({
      kind: X.literal("convert"),
      columns: cs,
      to: X.enum(["text", "number", "integer", "boolean", "date"]),
      decimal: X.enum([".", ","]),
      group: X.enum(["", ".", ",", " ", "'"]),
      dateOrder: X.enum(["MDY", "DMY"]),
    }),
    X.strictObject({
      kind: X.literal("fill"),
      columns: cs,
      direction: X.enum(["down", "up", "value"]),
      value: hg,
      groupColumns: X.array(fg).max(50),
    }),
    X.strictObject({ kind: X.literal("select"), columns: cs }),
    X.strictObject({
      kind: X.literal("pattern"),
      name: X.string().min(1).max(100),
      pattern: kV,
      examples: EV,
      inferenceColumns: cs.max(8),
    }),
  ]),
  CV = X.strictObject({
    sourceSheetId: X.string().min(1).max(100),
    sourceRange: X.string().min(1).max(50),
    header: X.boolean(),
    outputName: X.string().min(1).max(31),
    steps: X.array(SV).max(12),
  });
function mN(e) {
  let t = CV.safeParse(e);
  if (!t.success)
    throw Error(
      `Invalid field transformation: ${t.error.issues[0]?.message ?? "review the recipe"}`,
    );
  let n = t.data;
  if (
    ((n.outputName = n.outputName.trim()),
    !n.outputName ||
      /[\u0000-\u001f\u007f\\/\[\]*?:]/.test(n.outputName) ||
      /^'|'$/.test(n.outputName))
  )
    throw Error(
      "Choose a valid output worksheet name of 1\u201331 characters.",
    );
  for (let i of n.steps) {
    if (i.kind === "split") {
      if (
        i.mode === "delimited" &&
        (![...i.delimiters].length ||
          [...i.delimiters].length > 8 ||
          /["\r\n\0]/.test(i.delimiters))
      )
        throw Error(
          "Choose 1\u20138 delimiter characters without quotes, line breaks or NUL.",
        );
      if (
        i.mode === "fixed" &&
        (!i.breaks.length ||
          i.breaks.some((r, o) => o > 0 && r <= i.breaks[o - 1]))
      )
        throw Error(
          "Fixed-width breaks must be distinct increasing character positions.",
        );
    }
    if (i.kind === "convert" && i.to !== "text" && i.decimal === i.group)
      throw Error("Decimal and grouping separators must differ.");
    if (
      i.kind === "fill" &&
      new Set(i.groupColumns).size !== i.groupColumns.length
    )
      throw Error("Choose distinct fill group columns.");
    if (i.kind === "fill" && i.columns.some((r) => i.groupColumns.includes(r)))
      throw Error(
        "Fill targets must differ from the group columns. Use a separate preceding step to prepare group keys.",
      );
  }
  return n;
}
var pg = X.string().refine((e) => {
    try {
      let t = tt(e);
      return `${Ue(t.col)}${t.row + 1}` === e;
    } catch {
      return false;
    }
  }, "Invalid evidence address."),
  TV = X.strictObject({
    id: X.string().min(1).max(100),
    sourceId: X.string().min(1).max(100),
    sourceName: X.string().min(1).max(255),
    sourceSheet: X.string().min(1).max(31),
    originalRow: X.number().int().min(1).max(1048576),
    kind: X.enum(["file", "snapshot"]),
  }),
  AV = X.strictObject({
    address: pg,
    original: X.unknown().optional(),
    origin: TV.optional(),
  }),
  IV = X.strictObject({
    address: pg,
    row: X.number().int().min(1).max(1048576),
    field: X.string().min(1).max(150),
    value: hg,
    type: X.enum(["text", "number", "boolean", "date", "blank"]),
    sources: X.array(pg).min(1).max(500),
    exception: X.boolean(),
  }),
  RV = X.strictObject({
    step: X.number().int().min(1).max(12),
    row: X.number().int().min(1).max(1048576),
    field: X.string().min(1).max(150),
    raw: hg,
    sources: X.array(pg).min(1).max(500),
    reason: X.string().min(1).max(1e3),
  }),
  NV = X.strictObject({
    id: X.string().min(1).max(100),
    workbookId: X.string().min(1).max(100),
    sourceRevision: X.number().int().nonnegative(),
    sourceSheetId: X.string().min(1).max(100),
    sourceSheetName: X.string().min(1).max(31),
    outputSheetId: X.string().min(1).max(100),
    outputSheetName: X.string().min(1).max(31),
    createdAt: X.iso.datetime(),
    semantics: X.literal("reviewed-fields-v1"),
    request: X.unknown(),
    sources: X.array(AV).min(1).max(500),
    outputs: X.array(IV).min(1).max(5e3),
    exceptions: X.array(RV).max(5e3),
  });
function gN(e, t) {
  if (e === undefined) return;
  if (
    !Array.isArray(e) ||
    e.length > 100 ||
    JSON.stringify(e).length > fN.evidenceCharacters
  )
    throw Error(
      "Field transformation evidence exceeds its native history budget. Keep the original backup.",
    );
  let n = new Set(),
    i = 0;
  t.fieldTransforms = e.map((r) => {
    let o = NV.safeParse(r);
    if (!o.success)
      throw Error(
        `Invalid field transformation evidence: ${o.error.issues[0]?.message}`,
      );
    let a = o.data,
      s = mN(a.request),
      l = Qe(s.sourceRange);
    if (
      new Date(a.createdAt).toISOString() !== a.createdAt ||
      a.sources.length !==
        (l.lastRow - l.firstRow + 1) * (l.lastCol - l.firstCol + 1)
    )
      throw Error(
        "Invalid field evidence timestamp or incomplete original source population.",
      );
    if (n.has(a.id) || (i += a.outputs.length) > 2e4)
      throw Error("Duplicate or over-capacity field transformation evidence.");
    if (
      (n.add(a.id),
      a.sourceSheetId !== s.sourceSheetId || a.outputSheetName !== s.outputName)
    )
      throw Error("Field transformation evidence does not match its recipe.");
    let c = new Set(),
      u = a.sources.map((f) => {
        let p = tt(f.address);
        if (
          p.row < l.firstRow ||
          p.row > l.lastRow ||
          p.col < l.firstCol ||
          p.col > l.lastCol
        )
          throw Error(
            "An original field source is outside the recorded range.",
          );
        if (c.has(f.address)) throw Error("Duplicate original field source.");
        c.add(f.address);
        let h = f.original === undefined ? undefined : Oi(f.original);
        if (
          h &&
          typeof h.raw == "string" &&
          h.raw.startsWith("=") &&
          !h.literal
        )
          throw Error("Field source evidence must contain reviewed values.");
        return { ...f, ...(h ? { original: h } : {}) };
      }),
      d = new Set();
    for (let f of a.outputs) {
      if (d.has(f.address))
        throw Error("Duplicate transformed output address.");
      if (
        (d.add(f.address),
        ((f.type === "number" || f.type === "date") &&
          typeof f.value != "number") ||
          (f.type === "boolean" && typeof f.value != "boolean") ||
          (f.type === "text" && typeof f.value != "string") ||
          (f.type === "blank" && f.value !== null))
      )
        throw Error("Transformed output type does not match its value.");
    }
    for (let f of [...a.outputs, ...a.exceptions])
      if (
        f.row < l.firstRow + 1 + Number(s.header) ||
        f.row > l.lastRow + 1 ||
        f.sources.some((p) => !c.has(p)) ||
        new Set(f.sources).size !== f.sources.length
      )
        throw Error("Unknown or repeated original field source locator.");
    return { ...a, request: s, sources: u };
  });
}
var MV = new WeakSet(),
  $V = new WeakSet(),
  bN = new WeakMap();
var PV = new Set(["general", "number", "currency", "percent", "date"]),
  DV = (e, t) => Object.prototype.hasOwnProperty.call(e, t),
  Gs = (e) => !!e && typeof e == "object" && !Array.isArray(e),
  jt = (e) => {
    throw new Error(e);
  },
  Fi = (e, t, n = 32767) =>
    typeof e == "string" && e.length <= n
      ? e
      : jt(`${t} must be text of at most ${n} characters.`),
  mg = (e, t, n, i) =>
    typeof e == "number" && Number.isInteger(e) && e >= t && e <= n
      ? e
      : jt(`${i} must be an integer between ${t} and ${n}.`);
function Ue(e) {
  mg(e, 0, Ne.maxCols - 1, "Column");
  let t = "";
  for (let n = e + 1; n > 0; n = Math.floor((n - 1) / 26))
    t = String.fromCharCode(65 + ((n - 1) % 26)) + t;
  return t;
}
function tt(e) {
  let t = /^\$?([A-Za-z]{1,3})\$?([1-9]\d{0,6})$/.exec(e);
  if (!t) return jt(`Invalid cell address: ${e}`);
  let n = 0;
  for (let r of t[1].toUpperCase()) n = n * 26 + r.charCodeAt(0) - 64;
  let i = Number(t[2]) - 1;
  return i >= Ne.maxRows || n > Ne.maxCols
    ? jt(`Cell ${e} exceeds the supported workbook bounds.`)
    : { row: i, col: n - 1 };
}
function Qe(e) {
  let t = e.trim().split(":");
  if (t.length > 2) return jt(`Invalid range: ${e}`);
  let n = tt(t[0]),
    i = tt(t[1] ?? t[0]);
  return {
    firstRow: Math.min(n.row, i.row),
    lastRow: Math.max(n.row, i.row),
    firstCol: Math.min(n.col, i.col),
    lastCol: Math.max(n.col, i.col),
  };
}
function yN(e, t = true) {
  let n = null;
  for (let [i, r] of Object.entries(e.cells)) {
    let o = r.raw !== null && r.raw !== "",
      a =
        t &&
        (r.numberFormat ||
          r.hyperlink ||
          r.note ||
          r.tickmark ||
          r.numFmt ||
          r.format ||
          Object.keys(r.style ?? {}).length);
    if (!o && !a) continue;
    let { row, col } = tt(i);
    n
      ? ((n.firstRow = Math.min(n.firstRow, row)),
        (n.lastRow = Math.max(n.lastRow, row)),
        (n.firstCol = Math.min(n.firstCol, col)),
        (n.lastCol = Math.max(n.lastCol, col)))
      : (n = { firstRow: row, lastRow: row, firstCol: col, lastCol: col });
  }
  return n;
}
function Gm(e, t = true) {
  let n = yN(e, t);
  if (!n) return "A1";
  let i = `${Ue(n.firstCol)}${n.firstRow + 1}`,
    r = `${Ue(n.lastCol)}${n.lastRow + 1}`;
  return i === r ? i : `${i}:${r}`;
}
function mi(e, t) {
  return e.sheets.find((n) => n.id === t) ?? jt(`Sheet not found: ${t}`);
}
function Oi(e) {
  return gg(e, false);
}
function gg(e, t) {
  if (t && Hn(e) && MV.has(e) && !e.hyperlink && !e.style?.binding) return e;
  if (!Gs(e) || !DV(e, "raw")) return jt("A cell must have a raw value.");
  let n = e.raw;
  if (
    n !== null &&
    typeof n != "string" &&
    typeof n != "boolean" &&
    (typeof n != "number" || !Number.isFinite(n))
  )
    return jt("Cell values must be text, finite numbers, booleans, or null.");
  typeof n == "string" &&
    Fi(
      n,
      "Cell value",
      n.startsWith("=") && e.literal !== true
        ? dr.maxFormulaLength
        : dr.maxCellTextLength,
    );
  let i = { raw: n };
  if (e.literal !== undefined) {
    if (typeof e.literal != "boolean")
      return jt("Cell literal flag must be boolean.");
    i.literal = e.literal;
  }
  if (
    (e.numberFormat !== undefined && (i.numberFormat = Vw(e.numberFormat)),
    e.numFmt !== undefined &&
      (i.numFmt = Fi(e.numFmt, "Excel number format", 200)),
    e.format !== undefined)
  ) {
    if (typeof e.format != "string" || !PV.has(e.format))
      return jt("Unsupported cell format.");
    i.format = e.format;
  }
  return (
    e.hyperlink !== undefined && (i.hyperlink = Pw(e.hyperlink)),
    e.note !== undefined && (i.note = Fi(e.note, "Cell note", 1e4)),
    e.tickmark !== undefined && (i.tickmark = Fi(e.tickmark, "Tickmark", 40)),
    e.style !== undefined && (i.style = As(e.style)),
    i
  );
}
function gi(e) {
  return LV(e, false);
}
function LV(e, t) {
  if (t && Hn(e) && $V.has(e)) return e;
  if (!Gs(e) || e.schemaVersion !== 1)
    return jt("Unsupported workbook schema. Expected schemaVersion 1.");
  let n = Fi(e.id, "Workbook ID", 100),
    i = Fi(e.name, "Workbook name", 200);
  if (!n || !i.trim()) return jt("Workbook ID and name are required.");
  let r = mg(e.revision, 0, Number.MAX_SAFE_INTEGER - 1, "Revision");
  if (!Array.isArray(e.sheets) || !e.sheets.length)
    return jt("A workbook must contain at least one sheet.");
  let o = 0,
    a = new Set(),
    s = new Set(),
    l = e.sheets.map((p) => {
      if (!Gs(p)) return jt("Invalid sheet.");
      let h = Fi(p.id, "Sheet ID", 100),
        m = Fi(p.name, "Sheet name", 31);
      if (
        !h ||
        !m.trim() ||
        /[\\/\[\]*?:]/.test(m) ||
        a.has(h) ||
        s.has(m.toLowerCase())
      )
        return jt(
          "Sheet IDs and names must be nonempty and unique; sheet names cannot contain Excel-reserved characters.",
        );
      (a.add(h), s.add(m.toLowerCase()));
      let g = mg(p.rowCount, 1, Ne.maxRows, "Row count"),
        b = mg(p.colCount, 1, Ne.maxCols, "Column count");
      if (!Gs(p.cells)) return jt("Sheet cells must be an address map.");
      let w = t && Hn(p) ? bN.get(p) : undefined;
      if (w !== undefined) {
        if (((o += w), o > Ne.maxCells))
          return jt("Workbook exceeds the 100,000 populated cell limit.");
        let O = { ...p };
        return (bN.set(O, w), O);
      }
      let _ = {};
      for (let [O, E] of Object.entries(p.cells)) {
        let R = tt(O);
        if (O !== `${Ue(R.col)}${R.row + 1}` || R.row >= g || R.col >= b)
          return jt(`Invalid or out-of-bounds cell: ${O}`);
        if (++o > Ne.maxCells)
          return jt("Workbook exceeds the 100,000 populated cell limit.");
        _[O] = gg(E, t);
      }
      let S = {
        id: h,
        name: m,
        rowCount: Ne.maxRows,
        colCount: Ne.maxCols,
        cells: _,
      };
      if (p.columnWidths !== undefined) {
        if (!Gs(p.columnWidths))
          return jt("Column widths must be an index map.");
        S.columnWidths = {};
        for (let [O, E] of Object.entries(p.columnWidths)) {
          if (
            !/^(0|[1-9]\d*)$/.test(O) ||
            Number(O) >= b ||
            typeof E != "number" ||
            !Number.isFinite(E) ||
            E < 40 ||
            E > 600
          )
            return jt(
              "Column widths must be 40\u2013600 pixels for an existing column.",
            );
          S.columnWidths[O] = E;
        }
      }
      if (p.rowHeights !== undefined) {
        if (
          !Gs(p.rowHeights) ||
          (Object.getPrototypeOf(p.rowHeights) !== Object.prototype &&
            Object.getPrototypeOf(p.rowHeights) !== null)
        )
          return jt("Row heights must be a plain index map.");
        S.rowHeights = {};
        for (let [O, E] of Object.entries(p.rowHeights)) {
          if (
            !/^(0|[1-9]\d*)$/.test(O) ||
            Number(O) >= g ||
            typeof E != "number" ||
            !Number.isFinite(E) ||
            E < 21 ||
            E > 546
          )
            return jt(
              `Row heights must be ${21}\u2013${546} pixels for an existing row.`,
            );
          S.rowHeights[O] = E;
        }
      }
      if (p.hidden !== undefined) {
        if (typeof p.hidden != "boolean")
          return jt("Sheet hidden flag must be boolean.");
        S.hidden = p.hidden;
      }
      return (
        Rp(p, S),
        K6(p, S, { rows: g, cols: b }),
        Dw(p, S, { rows: g, cols: b }),
        E5(p, S),
        YR(p.delimitedSource, S),
        j2(p, S),
        p.objects !== undefined && Fd(p.objects, S),
        p.pageLayout !== undefined && Wd(p.pageLayout, S),
        S
      );
    });
  if (l.every((p) => p.hidden))
    return jt("At least one sheet must be visible.");
  if (!Array.isArray(e.events) || e.events.length > 2e3)
    return jt("Workbook history must be an array of at most 2,000 events.");
  let c = new Set(),
    u = 0,
    d = e.events.map((p) => {
      if (!Gs(p)) return jt("Invalid history event.");
      let h = Fi(p.id, "Event ID", 100),
        m = Fi(p.timestamp, "Event timestamp", 40);
      return !h || c.has(h) || !Number.isFinite(Date.parse(m))
        ? jt("History events require unique IDs and valid timestamps.")
        : (c.add(h),
          !Array.isArray(p.changes) ||
          p.changes.length > Ne.maxPatches ||
          (u += p.changes.length) > Ne.maxCells
            ? jt("Invalid history change count.")
            : {
                id: h,
                timestamp: m,
                actor: Fi(p.actor, "Actor", 100),
                action: Fi(p.action, "Action", 200),
                changes: p.changes.map((g) => {
                  if (
                    !Gs(g) ||
                    typeof g.sheetId != "string" ||
                    !g.sheetId.length ||
                    g.sheetId.length > 100 ||
                    typeof g.address != "string"
                  )
                    return jt("Invalid history cell reference.");
                  let b = tt(g.address);
                  return g.address !== `${Ue(b.col)}${b.row + 1}`
                    ? jt("History cell address is not canonical.")
                    : g.before === undefined && g.after === undefined
                      ? jt("History change must have a before or after value.")
                      : {
                          sheetId: g.sheetId,
                          address: g.address,
                          ...(g.before !== undefined
                            ? { before: gg(g.before, t) }
                            : {}),
                          ...(g.after !== undefined
                            ? { after: gg(g.after, t) }
                            : {}),
                        };
                }),
              });
    }),
    f = { schemaVersion: 1, id: n, name: i, revision: r, sheets: l, events: d };
  return (
    rR(e, f),
    e.datePolicy !== undefined && (f.datePolicy = go(e.datePolicy)),
    gR(e.savedFilterViews, f),
    y6(e, f),
    wR(e, f),
    _R(e, f),
    DR(e, f),
    CR(e, f),
    OR(e, f),
    M5(e, f),
    JR(e, f),
    qR(e, f),
    ZR(e, f),
    WR(e, f),
    G5(e, f),
    dN(e.textCleanups, f),
    gN(e.fieldTransforms, f),
    zm(e.tableEvidence, f),
    e.whatIfAnalyses !== undefined && (f.whatIfAnalyses = oN(e.whatIfAnalyses)),
    e.definedNames !== undefined && (f.definedNames = _d(e.definedNames, f)),
    _6(e, f),
    S5(e.structureProtection, f),
    Lw(e.metadata, f),
    B5(f),
    V6(e, f),
    z6(e.styleCatalog, f),
    e.formulaWatches !== undefined &&
      (f.formulaWatches = tN(e.formulaWatches, f)),
    dR(e, f),
    S6(e, f),
    f
  );
}
var wN = [
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
function so(e) {
  throw Error(e);
}
var vN = (e) => !!e && typeof e == "object" && !Array.isArray(e);
function xN(e) {
  if (e === undefined) return {};
  (!vN(e) ||
    Object.keys(e).some(
      (n) =>
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
        ].includes(n),
    )) &&
    so("Invalid or unknown chart settings.");
  let t = {};
  for (let n of ["altText", "xTitle", "yTitle", "units"])
    e[n] !== undefined &&
      ((typeof e[n] != "string" ||
        e[n].length > (n === "altText" ? 2e3 : 200)) &&
        so(`Invalid chart ${n}.`),
      (t[n] = e[n]));
  for (let n of ["xScale", "yScale"])
    e[n] !== undefined &&
      (["linear", "log"].includes(String(e[n])) ||
        so("Chart scales must be linear or log."),
      (t[n] = e[n]));
  if (
    (e.showLegend !== undefined &&
      (typeof e.showLegend != "boolean" && so("Invalid chart legend setting."),
      (t.showLegend = e.showLegend)),
    e.histogramBins !== undefined &&
      ((!Number.isInteger(e.histogramBins) ||
        Number(e.histogramBins) < 1 ||
        Number(e.histogramBins) > 50) &&
        so("Choose 1 to 50 histogram bins."),
      (t.histogramBins = Number(e.histogramBins))),
    e.sparklineKind !== undefined &&
      (["line", "column", "winLoss"].includes(String(e.sparklineKind)) ||
        so("Invalid sparkline style."),
      (t.sparklineKind = e.sparklineKind)),
    e.series !== undefined)
  ) {
    (!Array.isArray(e.series) || e.series.length > 20) &&
      so("Select at most 20 chart series.");
    let n = new Set();
    t.series = e.series.map((i) => {
      ((!vN(i) ||
        Object.keys(i).some((o) => !["column", "name", "kind"].includes(o)) ||
        !Number.isInteger(i.column) ||
        Number(i.column) < 1 ||
        Number(i.column) > 16384 ||
        n.has(Number(i.column))) &&
        so("Chart series require unique one-based worksheet columns."),
        n.add(Number(i.column)));
      let r = { column: Number(i.column) };
      return (
        i.name !== undefined &&
          ((typeof i.name != "string" ||
            !i.name.trim() ||
            i.name.length > 200) &&
            so("Chart series names require 1 to 200 characters."),
          (r.name = i.name)),
        i.kind !== undefined &&
          (["column", "line"].includes(String(i.kind)) ||
            so("Combo series must be column or line."),
          (r.kind = i.kind)),
        r
      );
    });
  }
  return (
    e.includedRows !== undefined &&
      ((!Array.isArray(e.includedRows) ||
        e.includedRows.length > 500 ||
        e.includedRows.some(
          (n) => !Number.isInteger(n) || n < 1 || n > 1048576,
        ) ||
        new Set(e.includedRows).size !== e.includedRows.length) &&
        so(
          "Chart filters require at most 500 unique one-based worksheet rows.",
        ),
      (t.includedRows = [...e.includedRows].sort((n, i) => n - i))),
    t
  );
}
var _N = 100,
  Xa = 2 * 1024 * 1024,
  UV = 500,
  mk = (e) => !!e && typeof e == "object" && !Array.isArray(e),
  bg = (e, t, n, i) => {
    if (typeof e != "number" || !Number.isFinite(e) || e < t || e > n)
      throw new Error(`Invalid drawing ${i}.`);
    return e;
  },
  Cc = (e, t, n) => {
    if (typeof e != "string" || e.length > t)
      throw new Error(`Invalid drawing ${n}.`);
    return e;
  },
  kN = (e) => {
    if (typeof e != "string" || !/^#[\da-f]{6}$/i.test(e))
      throw new Error("Drawing colors must be six-digit hex colors.");
    return e;
  };
function W2(e) {
  if (!(e instanceof Uint8Array) || e.length > Xa)
    throw Error("Picture hashing requires at most 2 MB of bytes.");
  return rr(e);
}
function Ha(e) {
  if (typeof e != "string" || e.length > Math.ceil(Xa / 3) * 4 + 30)
    throw Error("Pictures must be local PNG/JPEG data of at most 2 MB.");
  let t = /^data:image\/(png|jpeg);base64,([A-Za-z0-9+/]+={0,2})$/.exec(e);
  if (!t || t[2].length % 4 !== 0)
    throw Error("Pictures require canonical local PNG/JPEG base64 data.");
  let n;
  try {
    let a = atob(t[2]);
    if (btoa(a) !== t[2]) throw Error();
    n = Uint8Array.from(a, (s) => s.charCodeAt(0));
  } catch {
    throw Error("The picture data is invalid.");
  }
  if (n.length > Xa) throw Error("Pictures must be at most 2 MB.");
  let i = 0,
    r = 0,
    o = () => {
      if (!i || !r || i > 1e4 || r > 1e4 || i * r > 4e7)
        throw Error(
          "Picture dimensions are invalid or exceed 40 megapixels / 10,000 pixels per side.",
        );
    };
  if (t[1] === "png") {
    if (
      n.length < 45 ||
      ![137, 80, 78, 71, 13, 10, 26, 10].every((f, p) => n[p] === f)
    )
      throw Error("The file contents do not match a PNG picture.");
    let a = new DataView(n.buffer, n.byteOffset, n.byteLength);
    if (
      a.getUint32(8) !== 13 ||
      String.fromCharCode(...n.subarray(12, 16)) !== "IHDR"
    )
      throw Error("PNG requires an intact first IHDR chunk.");
    if (
      ((i = a.getUint32(16)),
      (r = a.getUint32(20)),
      o(),
      !{
        0: [1, 2, 4, 8, 16],
        2: [8, 16],
        3: [1, 2, 4, 8],
        4: [8, 16],
        6: [8, 16],
      }[n[25]]?.includes(n[24]) ||
        n[26] !== 0 ||
        n[27] !== 0 ||
        n[28] > 1)
    )
      throw Error("Unsupported PNG pixel encoding.");
    let l = 8,
      c = 0,
      u = false,
      d = false;
    for (; l < n.length;) {
      if (++c > 4096 || l + 12 > n.length)
        throw Error("PNG chunk structure is invalid or excessive.");
      let f = a.getUint32(l),
        p = String.fromCharCode(...n.subarray(l + 4, l + 8));
      if (f > n.length - l - 12 || !/^[A-Za-z]{4}$/.test(p))
        throw Error("PNG chunk length or type is invalid.");
      if (p === "IHDR" && l !== 8) throw Error("PNG has duplicate dimensions.");
      if ((p === "IDAT" && f > 0 && (u = true), p === "IEND")) {
        if (f !== 0 || l + 12 !== n.length)
          throw Error("PNG has invalid trailing data.");
        d = true;
      }
      l += f + 12;
    }
    if (!u || !d) throw Error("PNG data or end marker is missing.");
  } else {
    if (n.length < 12 || n[0] !== 255 || n[1] !== 216)
      throw Error("The file contents do not match a JPEG picture.");
    let a = 2,
      s = 0,
      l = false,
      c = false;
    for (; a < n.length;) {
      if (++s > 65536) throw Error("JPEG has too many segments.");
      if (l) for (; a < n.length && n[a] !== 255;) a++;
      if (a >= n.length || n[a++] !== 255)
        throw Error("JPEG marker structure is invalid.");
      for (; n[a] === 255;) a++;
      let u = n[a++];
      if (l && (u === 0 || (u >= 208 && u <= 215))) continue;
      if (u === 217) {
        if (a !== n.length) throw Error("JPEG has trailing data.");
        c = true;
        break;
      }
      if (u === 1) continue;
      if (a + 2 > n.length) throw Error("JPEG segment is truncated.");
      let d = n[a] * 256 + n[a + 1];
      if (d < 2 || a + d > n.length)
        throw Error("JPEG segment length is invalid.");
      if (u >= 192 && u <= 207 && ![196, 200, 204].includes(u)) {
        if (![192, 193, 194].includes(u) || d < 8 || n[a + 2] !== 8)
          throw Error("Unsupported JPEG pixel encoding.");
        let f = n[a + 7];
        if (f < 1 || f > 4 || d !== 8 + f * 3)
          throw Error("JPEG frame components are invalid.");
        let p = n[a + 3] * 256 + n[a + 4],
          h = n[a + 5] * 256 + n[a + 6];
        if (i && (i !== h || r !== p)) throw Error("JPEG dimensions conflict.");
        ((i = h), (r = p), o());
      }
      if (u === 218) {
        if (!i || d < 6 || d !== 6 + n[a + 2] * 2)
          throw Error("JPEG scan is invalid.");
        l = true;
      }
      a += d;
    }
    if (!l || !c) throw Error("JPEG scan or end marker is missing.");
    o();
  }
  return { bytes: n, extension: t[1], width: i, height: r };
}
function FV(e, t, n) {
  let i = Cc(e, t, n);
  if (!i.trim() || /[\u0000-\u001f\u007f]/.test(i))
    throw Error(
      `Invalid picture ${n}: use nonblank plain text without control characters.`,
    );
  return i;
}
function zV(e, t) {
  let { bytes } = Ha(e),
    i = W2(bytes);
  return EN(
    {
      ...t,
      version: 1,
      sourceReference: `sha256:${i}`,
      sha256: i,
      byteLength: bytes.length,
    },
    bytes,
  );
}
function EN(e, t) {
  if (
    !mk(e) ||
    ![Object.prototype, null].includes(Object.getPrototypeOf(e)) ||
    Reflect.ownKeys(e).some(
      (r) =>
        typeof r != "string" ||
        ![
          "version",
          "origin",
          "sourceReference",
          "sha256",
          "byteLength",
          "filename",
          "sourceLocator",
          "containerFilename",
          "mediaReference",
        ].includes(r),
    )
  )
    throw Error("Invalid picture evidence metadata.");
  if (
    e.version !== 1 ||
    typeof e.origin != "string" ||
    !["local-file", "xlsx-media", "legacy-native"].includes(e.origin)
  )
    throw Error("Invalid picture evidence origin.");
  let n = W2(t);
  if (
    typeof e.sha256 != "string" ||
    !/^[a-f0-9]{64}$/.test(e.sha256) ||
    e.sha256 !== n ||
    e.sourceReference !== `sha256:${n}` ||
    e.byteLength !== t.length
  )
    throw Error(
      "Picture evidence digest or byte length does not match the embedded image.",
    );
  let i = {
    version: 1,
    origin: e.origin,
    sourceReference: `sha256:${n}`,
    sha256: n,
    byteLength: t.length,
  };
  for (let r of [
    "filename",
    "sourceLocator",
    "containerFilename",
    "mediaReference",
  ])
    e[r] !== undefined &&
      (i[r] = FV(e[r], r === "sourceLocator" ? 2e3 : 512, r));
  if (i.origin === "local-file" && !i.filename)
    throw Error("A local evidence picture requires its original filename.");
  if (i.origin !== "xlsx-media" && (i.containerFilename || i.mediaReference))
    throw Error("Only observed XLSX media may have container metadata.");
  if (i.origin === "legacy-native" && i.filename)
    throw Error(
      "An original filename cannot be inferred for a legacy picture.",
    );
  return i;
}
function BV(e, t) {
  let n = Qe(e);
  if (
    n.lastRow >= t.rowCount ||
    n.lastCol >= t.colCount ||
    (n.lastRow - n.firstRow + 1) * (n.lastCol - n.firstCol + 1) > UV
  )
    throw new Error(
      "Select at most 500 cells inside this worksheet for a chart.",
    );
  return n;
}
function Fd(e, t) {
  if (e === undefined) return;
  if (!Array.isArray(e) || e.length > _N)
    throw new Error("A worksheet supports at most 100 drawing objects.");
  let n = new Set(),
    i = 0,
    r = 0;
  for (let a of e)
    if (mk(a) && a.kind === "picture" && typeof a.dataURL == "string") {
      if (a.dataURL.length > Math.ceil(Xa / 3) * 4 + 100)
        throw Error("Pictures must be at most 2 MB.");
      let s = a.dataURL.slice(a.dataURL.indexOf(",") + 1);
      if (
        ((r += Math.max(
          0,
          Math.floor((s.length * 3) / 4) -
            (s.endsWith("==") ? 2 : s.endsWith("=") ? 1 : 0),
        )),
        r > 10 * 1024 * 1024)
      )
        throw Error("Pictures in one worksheet may use at most 10 MB.");
    }
  let o = e.map((a) => {
    if (!mk(a) || !["picture", "shape", "chart"].includes(String(a.kind)))
      throw new Error("Invalid worksheet drawing object.");
    let s = Cc(a.id, 100, "id"),
      l = Cc(a.anchor, 20, "anchor");
    if (!s || n.has(s)) throw new Error("Drawing IDs must be unique.");
    n.add(s);
    let c = tt(l);
    if (c.row >= t.rowCount || c.col >= t.colCount)
      throw new Error("Drawing anchor exceeds the worksheet.");
    let u = {
      id: s,
      kind: a.kind,
      anchor: `${Ue(c.col)}${c.row + 1}`,
      offsetX: bg(a.offsetX, -4e3, 4e3, "horizontal offset"),
      offsetY: bg(a.offsetY, -4e3, 4e3, "vertical offset"),
      width: bg(a.width, 24, 4e3, "width"),
      height: bg(a.height, 24, 4e3, "height"),
      label: Cc(a.label, 500, "label"),
      fill: kN(a.fill),
      stroke: kN(a.stroke),
    };
    if (u.kind === "picture") {
      u.dataURL = Cc(a.dataURL, Math.ceil(Xa / 3) * 4 + 100, "picture");
      let d = Ha(u.dataURL);
      ((i += d.bytes.length),
        a.evidence !== undefined && (u.evidence = EN(a.evidence, d.bytes)));
    } else if (a.evidence !== undefined)
      throw Error("Evidence metadata belongs only to pictures.");
    if (u.kind === "shape") {
      if (!["rectangle", "ellipse", "arrow", "text"].includes(String(a.shape)))
        throw new Error("Unsupported shape.");
      u.shape = a.shape;
    }
    if (u.kind === "chart") {
      if (!wN.includes(a.chart)) throw new Error("Unsupported chart type.");
      if (
        ((u.chart = a.chart),
        (u.range = Cc(a.range, 50, "chart range")),
        BV(u.range, t),
        typeof a.headers != "boolean")
      )
        throw new Error("Invalid chart header setting.");
      ((u.headers = a.headers),
        a.chartOptions !== undefined && (u.chartOptions = xN(a.chartOptions)));
    }
    return u;
  });
  if (i > 10 * 1024 * 1024)
    throw new Error("Pictures in one worksheet may use at most 10 MB.");
  t.objects = o;
}
function jV(e, t) {
  let n = tt(t.split(":")[0]);
  return {
    id: crypto.randomUUID(),
    kind: e,
    anchor: `${Ue(n.col)}${n.row + 1}`,
    offsetX: 12,
    offsetY: 12,
    width: e === "chart" ? 480 : 240,
    height: e === "chart" ? 300 : 140,
    label: e === "picture" ? "Picture" : e === "chart" ? "Chart" : "Text",
    fill: "#E2EFDA",
    stroke: "#548235",
    ...(e === "shape" ? { shape: "rectangle" } : {}),
  };
}
var WV = (e) => {
  let t = "";
  for (let n of e) t += String.fromCharCode(n);
  return btoa(t);
};
function SN(e, t, n, i = {}) {
  let r = [],
    o = [],
    a = 0,
    s = 0;
  for (let l of t.getImages()) {
    if (++s > 1e3 || r.length >= _N) {
      o.push(
        "Additional pictures beyond the 100-object / 1,000-source inspection limit were omitted.",
      );
      break;
    }
    try {
      let c = e.getImage(Number(l.imageId)),
        u =
          c.extension === "jpeg"
            ? "jpeg"
            : c.extension === "png"
              ? "png"
              : null;
      if (!u) throw new Error("Only PNG/JPEG images are supported.");
      if (c.buffer && c.buffer.byteLength > Xa)
        throw Error("Embedded image exceeds 2 MB.");
      let d = c.base64?.startsWith("data:")
          ? c.base64.slice(c.base64.indexOf(",") + 1)
          : c.base64,
        f =
          d !== undefined
            ? Math.floor((d.length * 3) / 4) -
              (d.endsWith("==") ? 2 : d.endsWith("=") ? 1 : 0)
            : (c.buffer?.byteLength ?? 0);
      if (f > Xa) throw Error("Embedded image exceeds 2 MB.");
      if (a + f > 10 * 1024 * 1024)
        throw Error(
          "Additional picture omitted: pictures in one worksheet may use at most 10 MB.",
        );
      let p = c.base64?.startsWith("data:")
          ? c.base64
          : c.base64
            ? `data:image/${u};base64,${c.base64}`
            : `data:image/${u};base64,${WV(new Uint8Array(c.buffer))}`,
        h = Ha(p),
        m = Math.floor(l.range.tl.col),
        g = Math.floor(l.range.tl.row),
        b = jV("picture", `${Ue(m)}${g + 1}`);
      ((b.offsetX = Math.max(
        0,
        (l.range.tl.col - m) * (n.columnWidths?.[m] ?? 84),
      )),
        (b.offsetY = Math.max(
          0,
          (l.range.tl.row - g) * (n.rowHeights?.[g] ?? 21),
        )),
        (b.width = Math.max(24, Math.min(4e3, l.range.ext?.width ?? 240))),
        (b.height = Math.max(24, Math.min(4e3, l.range.ext?.height ?? 140))),
        (b.dataURL = p),
        (b.evidence = zV(p, {
          origin: "xlsx-media",
          ...(i.containerFilename !== undefined
            ? { containerFilename: i.containerFilename }
            : {}),
          mediaReference: `imageId:${l.imageId}`,
        })),
        (b.label = `Picture ${r.length + 1}`),
        r.push(b),
        (a += h.bytes.length));
    } catch (c) {
      o.push(
        c instanceof Error ? c.message : "An image could not be imported.",
      );
    }
  }
  return (Fd(r, n), o);
}
var VV = new WeakMap();
function CN(e, t, n) {
  if (
    (zd(e), e.fileName !== n || e.byteLength !== t.length || e.sha256 !== rr(t))
  )
    throw Error("Source inspection does not match its original.");
  VV.set(e, { signature: JSON.stringify(e), bytes: t.slice() });
}
var TN = (e) => e.replaceAll("~", "~0").replaceAll("/", "~1"),
  yg = (e) => !!e && typeof e == "object" && !Array.isArray(e);
function qV(e) {
  let t = 0,
    n = 0,
    i = () => {
      for (; /[\t\n\r ]/.test(e[t] ?? "x");) t++;
    },
    r = () => {
      let s = t++;
      for (; t < e.length;) {
        let l = e[t++];
        if (l === '"') return JSON.parse(e.slice(s, t));
        l === "\\" && t++;
      }
      throw Error("Invalid JSON string.");
    };
  function o(s) {
    if (s > an.depth || ++n > an.nodes)
      throw Error("JSON nesting or node budget exceeded.");
    i();
    let l = e[t];
    if (l === '"') return r();
    if (l === "[") {
      (t++, i());
      let d = [];
      if (e[t] === "]") return (t++, d);
      for (;;) {
        if ((d.push(o(s + 1)), i(), e[t++] === "]")) return d;
        if (e[t - 1] !== ",") throw Error("Invalid JSON array.");
      }
    }
    if (l === "{") {
      (t++, i());
      let d = Object.create(null);
      if (e[t] === "}") return (t++, d);
      for (;;) {
        if ((i(), e[t] !== '"')) throw Error("Invalid JSON object.");
        let f = r();
        if (
          ["__proto__", "prototype", "constructor"].includes(f) ||
          Object.hasOwn(d, f)
        )
          throw Error("Unsafe or duplicate JSON field.");
        if ((i(), e[t++] !== ":")) throw Error("Invalid JSON field.");
        if (((d[f] = o(s + 1)), i(), e[t++] === "}")) return d;
        if (e[t - 1] !== ",") throw Error("Invalid JSON object.");
      }
    }
    for (let [d, f] of [
      ["true", true],
      ["false", false],
      ["null", null],
    ])
      if (e.startsWith(d, t)) return ((t += d.length), f);
    let c = /-?(?:0|[1-9]\d*)(?:\.\d+)?(?:[eE][+-]?\d+)?/.exec(e.slice(t));
    if (!c || c.index !== 0) throw Error("Invalid JSON value.");
    t += c[0].length;
    let u = Number(c[0]);
    if (
      !Number.isFinite(u) ||
      (Number.isInteger(u) && !Number.isSafeInteger(u))
    )
      throw Error("Unsafe JSON number; encode large identifiers as strings.");
    return u;
  }
  let a = o(0);
  if ((i(), t !== e.length)) throw Error("Trailing JSON content.");
  return a;
}
function XV(e, t) {
  return {
    value: typeof e == "object" && e !== null ? JSON.stringify(e) : e,
    type:
      typeof e == "object" && e !== null
        ? Array.isArray(e)
          ? "array"
          : "object"
        : e === null
          ? "null"
          : typeof e,
    path: t,
    confidence: 1,
  };
}
function HV(e) {
  let t = [],
    n = (i, r) => {
      if (Array.isArray(i)) i.length && t.push({ value: i, path: r });
      else if (yg(i))
        for (let [o, a] of Object.entries(i)) n(a, r + "/" + TN(o));
    };
  if (
    (n(e, ""),
    !t.length && yg(e) && t.push({ value: [e], path: "" }),
    !t.length)
  )
    throw Error("Choose a JSON object or a nonempty array.");
  if (t.length > an.tables) throw Error("Source table count exceeded.");
  return t.map((i, r) => {
    let o = yg(e) && i.value[0] === e,
      a = new Set(),
      s = i.value.map((c, u) => {
        let d = new Map(),
          f = o ? i.path : i.path + "/" + u,
          p = (h, m) => {
            if (yg(h) && Object.keys(h).length)
              for (let [g, b] of Object.entries(h)) p(b, m + "/" + TN(g));
            else (a.add(m), d.set(m, XV(h, f + m)));
          };
        return (p(c, ""), d);
      }),
      l = [...a];
    return {
      id: "json-" + r,
      name: i.path || "Root",
      path: i.path,
      ambiguous: t.length > 1,
      columns: l.map((c) => ({ name: c || "Value", path: c })),
      rows: s.map((c, u) =>
        l.map(
          (d) =>
            c.get(d) ?? {
              value: null,
              type: "null",
              path: i.path + (o ? "" : "/" + u) + d,
              confidence: 1,
              missing: true,
            },
        ),
      ),
    };
  });
}
function GV(e) {
  if (/<!DOCTYPE|<!ENTITY/i.test(e))
    throw Error("XML document types and entities are not supported.");
  if (/<\?xml[^>]*encoding\s*=\s*['"](?!utf-8['"])/i.test(e))
    throw Error("Use UTF-8 XML.");
  let t = [],
    n = [],
    i = 0,
    r = new AN.SaxesParser({ xmlns: true }),
    o = (u, d) => (u ? `{${u}}${d}` : d);
  r.on("opentag", (u) => {
    if (++i > an.nodes || t.length >= an.depth)
      throw Error("XML nesting or node budget exceeded.");
    let d = o(u.uri, u.local),
      f = t.at(-1),
      p = f?.children ?? n,
      h = {
        namespaces: { ...(f?.namespaces ?? {}), ...u.ns },
        name: d,
        path:
          (f?.path ?? "") +
          "/" +
          d +
          "[" +
          (p.filter((m) => m.name === d).length + 1) +
          "]",
        attributes: [],
        text: "",
        children: [],
      };
    for (let m of Object.values(u.attributes))
      if (m.uri !== "http://www.w3.org/2000/xmlns/")
        if (
          m.uri === "http://www.w3.org/2001/XMLSchema-instance" &&
          m.local === "type"
        ) {
          let [g, b] = m.value.includes(":")
            ? m.value.split(":")
            : ["", m.value];
          if (g && !h.namespaces[g])
            throw Error("XML type prefix is undeclared.");
          h.type = o(h.namespaces[g] ?? "", b);
        } else
          m.uri === "http://www.w3.org/2001/XMLSchema-instance" &&
          m.local === "nil"
            ? (h.nil = m.value)
            : h.attributes.push({
                name: o(m.uri, m.local),
                value: m.value,
                uri: m.uri,
                local: m.local,
              });
    (p.push(h), t.push(h));
  });
  let a = (u) => {
    let d = t.at(-1);
    if (d && ((d.text += u), d.text.length > 32767))
      throw Error("XML cell text budget exceeded.");
  };
  if (
    (r.on("text", a),
    r.on("cdata", a),
    r.on("closetag", () => {
      t.pop();
    }),
    r.on("error", () => {
      throw Error("Malformed XML source.");
    }),
    r.write(e).close(),
    n.length !== 1)
  )
    throw Error("XML needs one root element.");
  let s = [],
    l = (u) => {
      let d = new Map();
      for (let f of u.children) d.set(f.name, [...(d.get(f.name) ?? []), f]);
      for (let [f, p] of d)
        p.length > 1 ? s.push({ path: u.path + "/" + f, nodes: p }) : l(p[0]);
    };
  if (
    (l(n[0]),
    s.length || s.push({ path: n[0].path, nodes: [n[0]] }),
    s.length > an.tables)
  )
    throw Error("XML table count exceeded.");
  let c = (u) => {
    let d = u.text,
      f = "string",
      p = u.text,
      h = u.type?.replace("{http://www.w3.org/2001/XMLSchema}", ""),
      m = u.text.trim();
    if (u.nil !== undefined) {
      if (!["true", "1", "false", "0"].includes(u.nil))
        throw Error("Invalid XML nil value.");
      if (["true", "1"].includes(u.nil)) {
        if (m || u.children.length)
          throw Error("A nil XML field cannot contain a value.");
        ((d = null), (f = "null"));
      }
    }
    if (
      f !== "null" &&
      u.type?.startsWith("{http://www.w3.org/2001/XMLSchema}")
    ) {
      if (h === "boolean") {
        if (!["true", "1", "false", "0"].includes(m))
          throw Error("Invalid typed XML boolean.");
        ((d = m === "true" || m === "1"), (f = "boolean"));
      }
      if (
        [
          "decimal",
          "double",
          "float",
          "integer",
          "int",
          "long",
          "short",
          "byte",
          "nonNegativeInteger",
          "positiveInteger",
        ].includes(h)
      ) {
        let g = !["decimal", "double", "float"].includes(h),
          b = g
            ? /^[+-]?\d+$/
            : h === "decimal"
              ? /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)$/
              : /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?$/;
        if (
          ((d = Number(m)),
          !b.test(m) ||
            !Number.isFinite(d) ||
            (Number.isInteger(d) && !Number.isSafeInteger(d)) ||
            (h === "nonNegativeInteger" && d < 0) ||
            (h === "positiveInteger" && d <= 0))
        )
          throw Error("Invalid or unsafe typed XML number.");
        let w = {
          int: [-2147483648, 2147483647],
          short: [-32768, 32767],
          byte: [-128, 127],
        };
        if (w[h] && (d < w[h][0] || d > w[h][1]))
          throw Error("Typed XML integer is out of range.");
        f = "number";
      }
    }
    return {
      value: d,
      type: f,
      path: u.path,
      lexical: p,
      ...(u.type ? { declaredType: u.type } : {}),
      confidence: 1,
    };
  };
  return s.map((u, d) => {
    let f = new Set(),
      p = u.nodes.map((m) => {
        let g = new Map(),
          b = (_, S) => {
            (f.add(_), g.set(_, S));
          },
          w = (_, S) => {
            for (let O of _.attributes)
              b(S + "/@" + O.name, {
                value: O.value,
                type: "string",
                path: _.path + "/@" + O.name,
                confidence: 1,
              });
            if (_.children.length) {
              if (_.text.trim())
                throw Error(
                  "Mixed-content XML tables need explicit extraction.",
                );
              let O = new Map();
              for (let E of _.children) {
                let R = (O.get(E.name) ?? 0) + 1;
                (O.set(E.name, R),
                  w(
                    E,
                    S +
                      "/" +
                      E.name +
                      (_.children.filter((I) => I.name === E.name).length > 1
                        ? "[" + R + "]"
                        : ""),
                  ));
              }
            } else b(S, c(_));
          };
        return (w(m, ""), g);
      }),
      h = [...f];
    return {
      id: "xml-" + d,
      name: u.path.slice(-200),
      path: u.path,
      ambiguous: s.length > 1,
      columns: h.map((m) => ({ name: (m || "Value").slice(-200), path: m })),
      rows: p.map((m, g) =>
        h.map(
          (b) =>
            m.get(b) ?? {
              value: null,
              type: "null",
              path: u.nodes[g].path + b,
              confidence: 1,
              missing: true,
            },
        ),
      ),
    };
  });
}
function gk(e, t) {
  if (e.length < 1 || e.length > an.bytes)
    throw Error("Source byte budget exceeded.");
  let n = t.split(".").at(-1)?.toLowerCase(),
    i = n === "jpg" ? "jpeg" : n === "htm" ? "html" : n,
    r = [],
    o = "document-adapter",
    a = "Host extraction required",
    s = [];
  if (i === "json" || i === "xml") {
    let c = new TextDecoder("utf-8", { fatal: true })
      .decode(e)
      .replace(/^\uFEFF/, "");
    ((r = i === "json" ? HV(qV(c)) : GV(c)),
      (o = i === "json" ? "native-json" : "native-xml"),
      (a = "Ledgerglass structured source v1"),
      s.push(
        "String identifiers and untyped XML remain text. Nested JSON arrays/empty objects are retained as JSON values. Imported formula-like strings remain literal.",
      ));
  } else if (i === "pdf") {
    if (
      new TextDecoder()
        .decode(e.subarray(0, 8))
        .match(/^%PDF-1\.[0-7]|^%PDF-2\.0/) == null
    )
      throw Error("Source extension does not match PDF content.");
    s.push(
      "PDF bytes are retained without rendering or executing embedded content. A host extractor or matching extraction JSON is required.",
    );
  } else if (i === "png" || i === "jpeg") {
    let c = "";
    for (let u = 0; u < e.length; u += 8192)
      c += String.fromCharCode(...e.subarray(u, u + 8192));
    (Ha(`data:image/${i};base64,${btoa(c)}`),
      s.push(
        "Image extraction requires a host OCR adapter or matching extraction JSON.",
      ));
  } else if (i === "html") {
    let c = new TextDecoder("utf-8", { fatal: true }).decode(e);
    if (!/^\s*(?:<!doctype\s+html[^>]*>\s*)?<(?:html|table)\b/i.test(c))
      throw Error("Source extension does not match an HTML document.");
    s.push(
      "HTML is retained as inert bytes. No markup, scripts or resource links are rendered or fetched. A host table extractor or matching extraction JSON is required.",
    );
  } else throw Error("Choose JSON, XML, PDF, PNG, JPEG or HTML.");
  let l = zd({
    schemaVersion: 1,
    fileName: t,
    format: i,
    byteLength: e.length,
    sha256: rr(e),
    tables: r,
    method: o,
    extractor: a,
    warnings: s,
  });
  return (CN(l, e, t), l);
}
var LN = Bi(Mo(), 1);
var NN = Bi(Mo(), 1);
function IN(e, t, n) {
  if (
    n > dr.maxXlsxExpandedBytes ||
    (t > dr.xlsxRatioThresholdBytes &&
      t / Math.max(e, 1) > dr.maxXlsxExpansionRatio)
  )
    throw new Error(
      "The XLSX archive exceeds the safe expansion limit (50 MB or 200:1 ratio).",
    );
}
function RN(e, t, n = 0) {
  let i = (u, d) =>
      new RegExp(`(?:^|\\s)${d}\\s*=\\s*(["'])(.*?)\\1`).exec(u)?.[2],
    r = e.replace(/<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>/g, ""),
    o = 0,
    a = 0,
    s = -1,
    l = n,
    c = (u) => {
      throw new Error(
        `Worksheet ${t}: ${u} Import stopped; no cells were truncated.`,
      );
    };
  for (let u of r.matchAll(/<(\/?)(?:[\w.-]+:)?(dimension|row|c)\b([^>]*)>/g)) {
    let [, d, f, p] = u;
    if (d) {
      f === "row" && (o = 0);
      continue;
    }
    if (f === "dimension") {
      let h = i(p, "ref");
      h || c("The declared dimension has no valid cell range.");
      try {
        Qe(h);
      } catch {
        c(
          `Declared dimension ${h} exceeds or violates the supported worksheet bounds.`,
        );
      }
    } else if (f === "row") {
      let h = i(p, "r"),
        m = Number(h);
      ((!h || !/^[1-9]\d*$/.test(h) || m > dr.maxRows) &&
        c(`Row ${h ?? "(missing)"} must be between 1 and ${dr.maxRows}.`),
        (o || m <= a) && c("Rows must have distinct, ascending row numbers."),
        (a = m),
        (o = /\/\s*$/.test(p) ? 0 : m),
        (s = -1));
    } else {
      let h = i(p, "r");
      (!h || !o) &&
        c("Each cell must have an explicit address inside its source row.");
      let m;
      try {
        m = tt(h);
      } catch {
        c(`Cell ${h} exceeds or violates the supported worksheet bounds.`);
      }
      ((m.row + 1 !== o || m.col <= s) &&
        c(`Cell ${h} has a duplicate, unordered or mismatched row address.`),
        (s = m.col),
        ++l > dr.maxXlsxCellRecords &&
          c(
            `The XLSX exceeds ${dr.maxXlsxCellRecords.toLocaleString("en-US")} source cell records, including explicit blank cells.`,
          ));
    }
  }
  return l;
}
var ZV =
    "Open the original in Excel, repair it if needed, and save a clean, unencrypted XLSX copy.",
  St = (e) => {
    throw Error(`${e} ${ZV}`);
  },
  KV = new Uint32Array(256).map((e, t) => {
    for (let n = 0; n < 8; n++) t = t & 1 ? 3988292384 ^ (t >>> 1) : t >>> 1;
    return t >>> 0;
  });
function Tc(e, t) {
  for (let n of t) e = KV[(e ^ n) & 255] ^ (e >>> 8);
  return e;
}
function vg(e, t = {}) {
  let n = yk(e);
  return (
    (!n.has("[Content_Types].xml") ||
      (!n.has("xl/workbook.xml") &&
        !(t.allowBinaryWorkbookInventory && n.has("xl/workbook.bin")))) &&
      St("This archive is not a supported Excel workbook."),
    n
  );
}
function yk(e) {
  let t = new DataView(e.buffer, e.byteOffset, e.byteLength);
  (e.length >= 8 &&
    t.getUint32(0, true) === 3759263696 &&
    t.getUint32(4, true) === 3776623009 &&
    St(
      "Encrypted Office or legacy binary workbooks are unsupported. Remove the open password in the source application.",
    ),
    (e.length < 22 || t.getUint32(0, true) !== 67324752) &&
      St("This file is not a supported XLSX archive."));
  let n = -1;
  for (let f = e.length - 22; f >= Math.max(0, e.length - 65557); f--)
    if (
      t.getUint32(f, true) === 101010256 &&
      f + 22 + t.getUint16(f + 20, true) === e.length
    ) {
      n = f;
      break;
    }
  n < 0 && St("The XLSX archive is incomplete or has trailing data.");
  let i = t.getUint16(n + 10, true),
    r = t.getUint32(n + 12, true),
    o = t.getUint32(n + 16, true);
  (t.getUint16(n + 4, true) ||
    t.getUint16(n + 6, true) ||
    i !== t.getUint16(n + 8, true) ||
    i === 65535 ||
    i > dr.maxXlsxEntries ||
    o + r !== n) &&
    St("Multi-volume, ZIP64 or oversized XLSX archives are not supported.");
  let a = new Map(),
    s = new Set(),
    l = [],
    c = (f) => {
      try {
        return new TextDecoder("utf-8", { fatal: true }).decode(f);
      } catch {
        return St("Unsupported XLSX path encoding.");
      }
    },
    u = o,
    d = 0;
  for (let f = 0; f < i; f++) {
    (u + 46 > n || t.getUint32(u, true) !== 33639248) &&
      St("Invalid XLSX archive directory.");
    let p = t.getUint16(u + 8, true),
      h = t.getUint16(u + 10, true),
      m = t.getUint32(u + 16, true),
      g = t.getUint32(u + 20, true),
      b = t.getUint32(u + 24, true),
      w = t.getUint16(u + 28, true),
      _ = t.getUint16(u + 30, true),
      S = t.getUint16(u + 32, true),
      O = t.getUint32(u + 42, true);
    (p & 8257 && St("Encrypted XLSX archive entries are unsupported."),
      (p & -2063 ||
        ![0, 8].includes(h) ||
        [g, b, O].includes(4294967295) ||
        t.getUint16(u + 34, true)) &&
        St("Unsupported XLSX compression or ZIP64 metadata."),
      (u + 46 + w + _ + S > n ||
        O + 30 > o ||
        t.getUint32(O, true) !== 67324752) &&
        St("Invalid XLSX archive offsets."));
    let E = c(e.subarray(u + 46, u + 46 + w));
    ((!E ||
      s.has(E.toLowerCase()) ||
      /^[\/]/.test(E) ||
      /[\\\x00-\x1f\x7f%:]/.test(E) ||
      E.split("/").some(
        (A, L, y) => A === ".." || A === "." || (!A && L < y.length - 1),
      )) &&
      St("Unsafe or duplicate XLSX archive path."),
      s.add(E.toLowerCase()));
    let R = t.getUint16(O + 26, true),
      I = t.getUint16(O + 28, true),
      U = O + 30 + R + I,
      P = U + g;
    (P > o ||
      t.getUint16(O + 6, true) !== p ||
      t.getUint16(O + 8, true) !== h ||
      c(e.subarray(O + 30, O + 30 + R)) !== E) &&
      St("XLSX local headers do not match the archive directory.");
    let T = P;
    if (p & 8) {
      let A = P + (P + 4 <= o && t.getUint32(P, true) === 134695760 ? 4 : 0);
      ((A + 12 > o ||
        t.getUint32(A, true) !== m ||
        t.getUint32(A + 4, true) !== g ||
        t.getUint32(A + 8, true) !== b) &&
        St("Invalid XLSX streaming data descriptor."),
        (T = A + 12));
    } else
      (t.getUint32(O + 14, true) !== m ||
        t.getUint32(O + 18, true) !== g ||
        t.getUint32(O + 22, true) !== b) &&
        St("XLSX local sizes or checksum do not match the archive directory.");
    (h === 0 && g !== b && St("Invalid stored XLSX entry size."),
      l.push({ start: O, end: T }),
      (d += b),
      IN(g, b, d),
      a.set(E, { size: b, crc: m }),
      (u += 46 + w + _ + S));
  }
  return (
    l.sort((f, p) => f.start - p.start),
    (u !== n ||
      l.some((f, p) => f.start !== (p ? l[p - 1].end : 0)) ||
      l.at(-1)?.end !== o) &&
      St("Overlapping or unlisted XLSX archive entries."),
    a
  );
}
var JV =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
  YV = "http://schemas.openxmlformats.org/package/2006/relationships",
  bk = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/",
  QV = "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
  ON = () => ({
    warnings: new Set(),
    hasMainType: false,
    hasOfficeDocument: false,
    strings: 0,
    maxString: -1,
    styles: 0,
    maxStyle: -1,
    emptyStringCaches: new Map(),
  });
function Zs(e, t) {
  let n;
  try {
    n = decodeURIComponent(t);
  } catch {
    return St(`Invalid relationship target in ${e}.`);
  }
  /[\\\x00-\x1f\x7f:?#]/.test(n) && St(`Unsafe internal relationship in ${e}.`);
  let i =
      e === "_rels/.rels"
        ? []
        : e
            .replace(/_rels\/[^/]+\.rels$/, "")
            .split("/")
            .filter(Boolean),
    r = n.startsWith("/") ? [] : i;
  for (let o of n.split("/"))
    o === ".."
      ? (r.length || St(`Relationship escapes the XLSX package in ${e}.`),
        r.pop())
      : o && o !== "." && r.push(o);
  return r.join("/");
}
function MN(e, t, n, i) {
  let r = new NN.SaxesParser({ xmlns: true }),
    o = 0,
    a = 0,
    s = new Set(),
    l = [],
    c,
    u = /^xl\/worksheets\/[^/]+\.xml$/.test(t);
  (r.on("error", () => St(`Malformed XML in ${t}.`)),
    r.on("doctype", () =>
      St(`DTD/entity declarations are unsupported in ${t}.`),
    ),
    r.on("cdata", () =>
      St(
        `CDATA content is unsupported in ${t}. Resave as standard XLSX text encoding.`,
      ),
    ),
    r.on("opentag", (d) => {
      (++o > 128 || ++a > 1e6) &&
        St(`XML complexity exceeds the supported budget in ${t}.`);
      let f = (h) => d.attributes[h]?.value,
        p = l.at(-1);
      if (
        (l.push(d.local),
        t === "xl/sharedStrings.xml" && d.local === "si" && i.strings++,
        t === "xl/styles.xml" &&
          d.local === "xf" &&
          p === "cellXfs" &&
          i.styles++,
        u && d.local === "c" && p === "row")
      ) {
        let h = f("t") ?? "n";
        (["n", "s", "str", "inlineStr", "b", "e"].includes(h) ||
          St(`Unsupported cell type ${h} in ${t}.`),
          (c = {
            address: f("r") ?? "unknown cell",
            type: h,
            value: "",
            inValue: false,
            formula: false,
          }));
        let m = f("s");
        m !== undefined &&
          ((!/^\d+$/.test(m) || !Number.isSafeInteger(Number(m))) &&
            St(`Invalid cell style index in ${t}.`),
          (i.maxStyle = Math.max(i.maxStyle, Number(m))));
      }
      if (
        (c &&
          p === "c" &&
          d.local === "v" &&
          ((c.value = ""), (c.inValue = true)),
        c &&
          p === "c" &&
          d.local === "f" &&
          ((c.formula = true),
          f("t") &&
            !["normal", "shared", "array"].includes(f("t")) &&
            St(
              `Unsupported formula kind at ${c.address} in ${t}. Convert data-table formulas to reviewed ordinary formulas or values first.`,
            )),
        u &&
          ["sheetProtection", "autoFilter", "extLst", "hyperlinks"].includes(
            d.local,
          ) &&
          i.warnings.add(
            `Worksheet ${d.local} metadata was detected. Protection, filter criteria, interactive links and unsupported extensions are not preserved as editable Excel features.`,
          ),
        o === 1 &&
          ((t === "[Content_Types].xml" &&
            (d.local !== "Types" ||
              d.uri !==
                "http://schemas.openxmlformats.org/package/2006/content-types")) ||
            (t.endsWith(".rels") &&
              (d.local !== "Relationships" || d.uri !== YV)) ||
            (t === "xl/workbook.xml" &&
              (d.local !== "workbook" || d.uri !== QV))) &&
          St(
            `Unsupported XLSX package type or Strict XML namespace in ${t}. Save as standard Excel Workbook (.xlsx).`,
          ),
        t === "[Content_Types].xml" && d.local === "Override")
      ) {
        let h = f("PartName"),
          m = f("ContentType");
        ((!h || !m || s.has(h)) &&
          St("Invalid or duplicate XLSX content type declaration."),
          s.add(h),
          /model\+data|dataModel/i.test(m) &&
            i.warnings.add(
              "Imported Excel Data Model and DAX content is not executed or translated; model caches are historical and require independent review. Exact bytes remain only in the retained original.",
            ),
          /macroEnabled|vbaProject|activeX|oleObject/i.test(m) &&
            St(
              "Macro-enabled or executable workbook content is unsupported; renaming a file does not convert its format.",
            ),
          h === "/xl/workbook.xml" &&
            (m !== JV &&
              St(
                "Unsupported workbook content type. Templates and alternate formats must be saved as XLSX.",
              ),
            (i.hasMainType = true)),
          n.has(Zs("_rels/.rels", h)) || St(`Missing XLSX part ${h}.`));
      }
      if (t.endsWith(".rels") && d.local === "Relationship") {
        let h = f("Id"),
          m = f("Target"),
          g = f("Type"),
          b = f("TargetMode");
        if (
          ((!h ||
            !m ||
            !g ||
            s.has(h) ||
            (b && !["External", "Internal"].includes(b))) &&
            St(`Invalid or duplicate relationship in ${t}.`),
          s.add(h),
          /vbaProject|activeX|oleObject|attachedTemplate/i.test(g) &&
            St("Executable or embedded-object relationships are unsupported."),
          b === "External")
        )
          (g !== `${bk}hyperlink` &&
            g !== `${bk}externalLinkPath` &&
            St(`Unsupported external relationship in ${t}.`),
            i.warnings.add(
              "External hyperlinks/workbook links were found. No external target is fetched; interactive links and refresh connections are omitted.",
            ));
        else {
          let w = Zs(t, m);
          (n.has(w) || St(`Missing relationship target ${w} in ${t}.`),
            t === "_rels/.rels" &&
              g === `${bk}officeDocument` &&
              ((i.hasOfficeDocument || w !== "xl/workbook.xml") &&
                St("Unsupported or duplicate workbook relationship."),
              (i.hasOfficeDocument = true)));
        }
      }
    }),
    r.on("text", (d) => {
      c?.inValue && (c.value += d);
    }),
    r.on("closetag", (d) => {
      if (c && d.local === "v") {
        let f = c.value;
        if (
          (c.type === "n" &&
            f !== "" &&
            (!/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(f) ||
              !Number.isFinite(Number(f))) &&
            St(`Invalid numeric value at ${c.address} in ${t}.`),
          c.type === "b" &&
            !["0", "1"].includes(f) &&
            St(
              `Unsupported boolean value at ${c.address} in ${t}; save numeric Boolean encoding in a standard XLSX copy.`,
            ),
          c.type === "s" &&
            ((!/^\d+$/.test(f) || !Number.isSafeInteger(Number(f))) &&
              St(`Invalid shared-string index at ${c.address} in ${t}.`),
            (i.maxString = Math.max(i.maxString, Number(f)))),
          c.type === "e" &&
            !/^#(?:NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|SPILL!|CALC!|GETTING_DATA|CONNECT!|BLOCKED!|FIELD!|UNKNOWN!|BUSY!)$/.test(
              f,
            ) &&
            St(`Unsupported error value at ${c.address} in ${t}.`),
          c.type === "e" &&
            !c.formula &&
            i.warnings.add(
              "Source constant error cells are flattened to displayed text. Error propagation from those cells is not preserved; review dependent calculations.",
            ),
          c.type === "str" && c.formula && f === "")
        ) {
          let p = i.emptyStringCaches.get(t) ?? new Set();
          (p.add(c.address), i.emptyStringCaches.set(t, p));
        }
        c.inValue = false;
      }
      (d.local === "c" && (c = undefined), l.pop(), o--);
    }),
    r.write(e).close());
}
function $N(e, t) {
  ((!e.hasMainType || !e.hasOfficeDocument) &&
    St("Missing standard XLSX workbook content type or relationship."),
    e.maxString >= e.strings &&
      St("A worksheet references a missing shared string."),
    e.maxStyle >= e.styles &&
      St("A worksheet references a missing cell style."));
  let n = [
    [
      /^xl\/(?:model|dataModel)\//i,
      "Excel Data Model / Power Pivot content is retained only in the original source and is not executed or translated. Relationships, DAX measures and model caches are unavailable; worksheet results are historical snapshots. Rebuild an explicit local Data Model from reviewed worksheet tables before relying on recalculated results.",
    ],
    [/^xl\/charts\//i, "Charts are not imported as editable native charts."],
    [
      /^xl\/(?:pivotTables|pivotCache)\//i,
      "Pivot definitions and caches are omitted; supported worksheet result cells remain.",
    ],
    [
      /^xl\/(?:queryTables|externalLinks)\/|^xl\/connections\.xml$/i,
      "Queries, external workbook links and connections are omitted and never refreshed.",
    ],
    [
      /^customXml\/|^docProps\/custom\.xml$/i,
      "Custom XML and custom document properties are not editable; the part inventory identifies passive metadata preserved unchanged and content retained only in the original.",
    ],
    [
      /^xl\/(?:threadedComments|persons|slicers|slicerCaches|timelines)\//i,
      "Threaded discussions, people, slicers and timelines are omitted.",
    ],
  ];
  for (let i of t) for (let [r, o] of n) r.test(i) && e.warnings.add(o);
  return [...e.warnings];
}
function PN(e) {
  let t = () => {
      throw Error(
        "Unsupported, encrypted or corrupt legacy workbook. Open the original in its source application and save an unencrypted XLSX copy.",
      );
    },
    n = new DataView(e.buffer, e.byteOffset, e.byteLength),
    i = 4294967294,
    r = 4294967295;
  if (
    e.length < 512 ||
    n.getUint32(0, true) !== 3759263696 ||
    n.getUint32(4, true) !== 3776623009 ||
    n.getUint16(28, true) !== 65534
  )
    return t();
  let o = n.getUint16(26, true),
    a = n.getUint16(30, true),
    s = 2 ** a;
  if (
    !((o === 3 && a === 9) || (o === 4 && a === 12)) ||
    n.getUint16(32, true) !== 6 ||
    n.getUint32(56, true) !== 4096 ||
    e.length % s
  )
    return t();
  let l = e.length / s - 1,
    c = (A) =>
      !Number.isInteger(A) || A < 0 || A >= l
        ? t()
        : e.subarray((A + 1) * s, (A + 2) * s),
    u = [],
    d = new Set();
  for (let A = 0; A < 109; A++) {
    let L = n.getUint32(76 + A * 4, true);
    L !== r && u.push(L);
  }
  let f = n.getUint32(68, true),
    p = n.getUint32(72, true);
  if (p > l) return t();
  for (let A = 0; A < p; A++) {
    if (d.has(f)) return t();
    d.add(f);
    let L = c(f),
      y = new DataView(L.buffer, L.byteOffset, L.byteLength);
    for (let v = 0; v < s - 4; v += 4) {
      let k = y.getUint32(v, true);
      k !== r && u.push(k);
    }
    f = y.getUint32(s - 4, true);
  }
  if (
    (p && f !== i) ||
    u.length !== n.getUint32(44, true) ||
    !u.length ||
    u.length > l ||
    new Set(u).size !== u.length
  )
    return t();
  let h = [];
  for (let A of u) {
    let L = c(A),
      y = new DataView(L.buffer, L.byteOffset, L.byteLength);
    for (let v = 0; v < s; v += 4) h.push(y.getUint32(v, true));
  }
  for (let A of u) if (h[A] !== 4294967293) return t();
  for (let A of d) if (h[A] !== 4294967292) return t();
  let m = (A, L, y, v, k) => {
      let x = new Set(),
        C = [],
        N = A,
        F = 0;
      for (; N !== i;) {
        if (
          N >= L.length ||
          x.has(N) ||
          x.size >= v ||
          (u.includes(N) && L === h) ||
          (d.has(N) && L === h)
        )
          return t();
        x.add(N);
        let q = y(N);
        if ((C.push(q), (F += q.length), (N = L[N]), F > e.length)) return t();
      }
      if (k !== undefined && (k > F || F - k >= (C[0]?.length ?? 1)))
        return t();
      let z = new Uint8Array(k ?? F),
        j = 0;
      for (let q of C)
        (z.set(q.subarray(0, z.length - j), j),
          (j = Math.min(j + q.length, z.length)));
      return z;
    },
    g = m(n.getUint32(48, true), h, c, l),
    b = new DataView(g.buffer);
  if (!g.length || g.length % 128) return t();
  let w = [];
  for (let A = 0; A < g.length; A += 128) {
    let L = b.getUint8(A + 66),
      y = b.getUint16(A + 64, true),
      v = b.getUint32(A + 124, true);
    if (
      ![0, 1, 2, 5].includes(L) ||
      (L && (y < 2 || y > 64 || y % 2 || b.getUint16(A + y - 2, true)))
    )
      return t();
    let k = L
      ? new TextDecoder("utf-16le", { fatal: true }).decode(
          g.subarray(A, A + y - 2),
        )
      : "";
    if (
      L &&
      (!k || /[\x00-\x1f]/.test(k.replace(/^\x05/, "")) || (v && o === 4))
    )
      return t();
    w.push({
      name: k,
      type: L,
      left: b.getUint32(A + 68, true),
      right: b.getUint32(A + 72, true),
      child: b.getUint32(A + 76, true),
      start: b.getUint32(A + 116, true),
      size: b.getUint32(A + 120, true),
    });
  }
  if (
    w[0].type !== 5 ||
    w.some((A) => /^(EncryptedPackage|EncryptionInfo)$/i.test(A.name))
  )
    return t();
  let _ = [],
    S = new Set(),
    O = [w[0].child];
  for (; O.length;) {
    let A = O.pop();
    if (A === r) continue;
    if (!A || A >= w.length || S.has(A)) return t();
    S.add(A);
    let L = w[A];
    if (![1, 2].includes(L.type)) return t();
    (_.push(L), O.push(L.left, L.right));
  }
  let E = _.filter((A) => A.type === 2 && /^(Workbook|Book)$/i.test(A.name));
  if (E.length !== 1) return t();
  let R = E[0];
  if (R.size < 8 || R.size > e.length) return t();
  let I;
  if (R.size >= 4096) I = m(R.start, h, c, l, R.size);
  else {
    let A = n.getUint32(64, true);
    if (!A || A > l) return t();
    let L = m(n.getUint32(60, true), h, c, A, A * s),
      y = new DataView(L.buffer),
      v = [];
    for (let x = 0; x < L.length; x += 4) v.push(y.getUint32(x, true));
    let k = m(w[0].start, h, c, l, w[0].size);
    I = m(
      R.start,
      v,
      (x) => ((x + 1) * 64 > k.length ? t() : k.subarray(x * 64, (x + 1) * 64)),
      Math.ceil(k.length / 64),
      R.size,
    );
  }
  let U = new DataView(I.buffer),
    P = U.getUint16(4, true);
  if (
    U.getUint16(0, true) !== 2057 ||
    U.getUint16(2, true) < 4 ||
    ![1280, 1536].includes(P) ||
    U.getUint16(6, true) !== 5
  )
    return t();
  let T = false;
  for (let A = 0; A + 4 <= I.length;) {
    let L = U.getUint16(A, true),
      y = U.getUint16(A + 2, true);
    if (A + 4 + y > I.length || L === 47) return t();
    if (((A += 4 + y), L === 10)) {
      T = true;
      break;
    }
  }
  return T
    ? {
        paths: w.filter((A) => A.type).map((A) => A.name),
        workbookBytes: I.length,
        version: P === 1536 ? "BIFF8" : "BIFF5",
      }
    : t();
}
var eq = new WeakMap();
function DN(e, t, n) {
  if (
    !e ||
    e.schemaVersion !== 1 ||
    e.fileName !== n ||
    e.fileBytes !== t.length ||
    e.sha256 !== rr(t) ||
    !_c.includes(e.format) ||
    !n.toLowerCase().endsWith("." + e.format) ||
    e.execution !== "never" ||
    !Array.isArray(e.parts) ||
    e.parts.length > 2e3
  )
    throw Error("Invalid isolated format inspection.");
  eq.set(e, { signature: JSON.stringify(e), bytes: t.slice() });
}
var _ie = Object.freeze(
    _c.map((e) =>
      Object.freeze({
        format: e,
        read:
          e === "xltx"
            ? "Native template adapter: supported XLSX subset"
            : "External conversion to XLSX, then reviewed import",
        write:
          e === "xltx"
            ? "Derived XLTX template with loss review"
            : "Export reviewed XLSX, then Save As in the source application",
        preserve:
          "Exact original bytes in native JSON; derived output is not lossless",
        formulas:
          e === "xltx"
            ? "Supported formulas and cache comparison"
            : "Verify formulas and caches after external conversion",
        dates: "Verify date system, serials and displayed dates",
        styles: "Supported formatting only; compare with original",
        specialParts:
          "Original only unless supported by the XLSX codec; never execute macros or objects",
      }),
    ),
  ),
  wg = Object.freeze({
    xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
    xltx: "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml",
    xlsm: "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
    xltm: "application/vnd.ms-excel.template.macroEnabled.main+xml",
    xlsb: "application/vnd.ms-excel.sheet.binary.macroEnabled.main",
  });
function tq(e) {
  return [
    "Keep the original file. Inspection and conversion here never execute macros, objects or external refreshes.",
    e === "ods"
      ? "In a trusted source application that supports ODS, open the original and Save As an unencrypted Excel Workbook (.xlsx)."
      : "In a trusted source application that supports this format, open the original with macros and external refresh disabled. Save As an unencrypted Excel Workbook (.xlsx).",
    "Review the source application's compatibility warnings. Keep any needed automation, objects and unsupported metadata in the original.",
    "Choose that converted XLSX here. Compare formulas, date systems, dates, styles and special parts with the original; confirm the loss acknowledgement and review the XLSX compatibility report before replacing the workbook.",
    "For output in the original format, export a reviewed derived XLSX, open it in the source application and Save As the target format. Reopen there and verify the result. Renaming an extension does not convert a workbook.",
  ];
}
function vk(e) {
  if (e.length > Ne.maxFileBytes)
    throw Error("Files must be 10 MB or smaller.");
  let t = yk(e),
    n = Object.create(null),
    i = new Set(),
    r = 0,
    o,
    a = new gc((s) => {
      if (!t.has(s.name) || i.has(s.name)) {
        o = Error("Alternate workbook archive directory mismatch.");
        return;
      }
      i.add(s.name);
      let l = 0,
        c = -1,
        u = [];
      ((s.ondata = (d, f, p) => {
        if (!o) {
          if (d) {
            o = Error("Corrupt alternate workbook archive.");
            return;
          }
          if (
            ((l += f.length),
            (r += f.length),
            l > t.get(s.name).size || r > 50 * 1024 * 1024)
          ) {
            ((o = Error(
              "Alternate workbook archive exceeds its expansion budget.",
            )),
              s.terminate());
            return;
          }
          if (((c = Tc(c, f)), u.push(f), p)) {
            if (
              l !== t.get(s.name).size ||
              (c ^ -1) >>> 0 !== t.get(s.name).crc
            ) {
              o = Error("Alternate workbook archive checksum mismatch.");
              return;
            }
            let h = new Uint8Array(l),
              m = 0;
            for (let g of u) (h.set(g, m), (m += g.length));
            n[s.name] = h;
          }
        }
      }),
        s.start());
    });
  a.register(mc);
  try {
    a.push(e, true);
  } catch {
    throw Error("Unable to inspect corrupt alternate workbook archive.");
  }
  if (o) throw o;
  if (Object.keys(n).length !== t.size)
    throw Error("Alternate workbook archive is incomplete.");
  return n;
}
function Ac(e, t) {
  let n = new TextDecoder("utf-8", { fatal: true }).decode(e);
  if (
    n.includes("\0") ||
    /<\?xml[^>]*encoding\s*=\s*["'](?!utf-8|utf8)[^"']+/i.test(n)
  )
    throw Error("Alternate workbook XML must use UTF-8.");
  let i = new LN.SaxesParser({ xmlns: true }),
    r = 0,
    o = 0;
  (i.on("doctype", () => {
    throw Error("Alternate workbook XML declarations are unsupported.");
  }),
    i.on("opentag", (a) => {
      if (++r > 128 || ++o > 1e6)
        throw Error("Alternate workbook XML exceeds its inspection budget.");
      t(a);
    }),
    i.on("closetag", () => {
      r--;
    }),
    i.write(n).close());
}
function xg(e) {
  let t = (s) =>
    Object.fromEntries(
      Object.values(s.attributes).map((l) => [l.local, l.value]),
    );
  if (e.mimetype) {
    if (
      new TextDecoder().decode(e.mimetype) !==
        "application/vnd.oasis.opendocument.spreadsheet" ||
      !e["META-INF/manifest.xml"] ||
      !e["content.xml"] ||
      e["[Content_Types].xml"]
    )
      throw Error("Unsupported or mismatched OpenDocument workbook identity.");
    let s = false,
      l = false;
    if (
      (Ac(e["META-INF/manifest.xml"], (c) => {
        let u = t(c);
        if (c.uri === "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0") {
          if (c.local === "encryption-data")
            throw Error(
              "Encrypted OpenDocument workbooks require an unencrypted XLSX conversion.",
            );
          c.local === "file-entry" &&
            u["full-path"] === "/" &&
            u["media-type"] ===
              "application/vnd.oasis.opendocument.spreadsheet" &&
            (s = true);
        }
      }),
      Ac(e["content.xml"], (c) => {
        c.local === "spreadsheet" &&
          c.uri === "urn:oasis:names:tc:opendocument:xmlns:office:1.0" &&
          (l = true);
      }),
      !s || !l)
    )
      throw Error("Unsupported or mismatched OpenDocument workbook identity.");
    return "ods";
  }
  if (!e["[Content_Types].xml"] || !e["_rels/.rels"])
    throw Error("Unsupported alternate workbook package identity.");
  let n,
    i = "",
    r = 0,
    o = new Set();
  Ac(e["[Content_Types].xml"], (s) => {
    if (
      s.uri !== "http://schemas.openxmlformats.org/package/2006/content-types"
    )
      throw Error("Unsupported alternate workbook content-type namespace.");
    let l = t(s);
    if (s.local === "Override") {
      if (!l.PartName || o.has(l.PartName))
        throw Error("Duplicate or missing alternate workbook content type.");
      if (
        (o.add(l.PartName),
        l.PartName === "/xl/workbook.xml" || l.PartName === "/xl/workbook.bin")
      ) {
        if (n) throw Error("Ambiguous alternate workbook identity.");
        ((n = l.ContentType), (i = l.PartName.slice(1)));
      }
    }
  });
  let a = Object.entries(wg).find(([, s]) => s === n)?.[0];
  if (
    !a ||
    !e[i] ||
    i !== (a === "xlsb" ? "xl/workbook.bin" : "xl/workbook.xml")
  )
    throw Error("Unsupported alternate workbook content type.");
  if (
    (Ac(e["_rels/.rels"], (s) => {
      let l = t(s);
      if (
        s.uri ===
          "http://schemas.openxmlformats.org/package/2006/relationships" &&
        s.local === "Relationship" &&
        l.Type ===
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
      ) {
        if (l.TargetMode || Zs("_rels/.rels", l.Target) !== i)
          throw Error("Invalid alternate workbook main relationship.");
        r++;
      }
    }),
    a === "xlsb")
  ) {
    let s = e[i];
    if (s.length < 3 || s[0] !== 131 || s[1] !== 1 || s[2] !== 0)
      throw Error(
        "Invalid binary workbook beginning record. Convert the original in its source application.",
      );
  } else {
    let s = true;
    Ac(e[i], (l) => {
      if (
        s &&
        ((s = false),
        l.local !== "workbook" ||
          ![
            "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
            "http://purl.oclc.org/ooxml/spreadsheetml/main",
          ].includes(l.uri))
      )
        throw Error("Invalid alternate workbook XML identity.");
    });
  }
  if (r !== 1) throw Error("Invalid alternate workbook main relationship.");
  return a;
}
function Vd(e, t) {
  if (
    !t ||
    t.length > 255 ||
    /[\\/\x00-\x1f\x7f\u202a-\u202e\u2066-\u2069]/.test(t) ||
    !e.length ||
    e.length > Ne.maxFileBytes
  )
    throw Error(
      "Invalid alternate workbook filename or 10 MB file limit exceeded.",
    );
  let n = t.split(".").pop().toLowerCase();
  if (!_c.includes(n))
    throw Error("Choose an XLS, XLSB, XLSM, XLTX, XLTM or ODS source.");
  let i, r, o;
  if (e[0] === 208 && e[1] === 207) {
    let s = PN(e);
    ((i = "xls"),
      (r =
        s.version + " workbook in a CFB container; binary cells not decoded"),
      (o = s.paths.map((l) => ({
        path: l,
        bytes: /^(Workbook|Book)$/i.test(l) ? s.workbookBytes : 0,
        disposition: "original-only",
      }))));
  } else {
    let s = vk(e);
    ((i = xg(s)),
      (r =
        i === "ods"
          ? "ODF spreadsheet MIME, manifest and content"
          : "OOXML main content type and package relationship"));
    for (let [l, c] of Object.entries(s))
      /\.(xml|rels)$/i.test(l) && Ac(c, () => {});
    o = Object.entries(s).map(([l, c]) => ({
      path: l,
      bytes: c.length,
      sha256: rr(c),
      disposition: "original-only",
    }));
  }
  if (i !== n)
    throw Error(
      "The filename extension does not match the workbook content. Restore its original extension or convert it in the source application; renaming is not conversion.",
    );
  let a = {
    schemaVersion: 1,
    fileName: t,
    format: n,
    fileBytes: e.length,
    sha256: rr(e),
    identity: r,
    execution: "never",
    parts: o,
    steps: tq(n),
    warnings: [
      "Identity inspection is not workbook conversion, malware scanning or a guarantee of formula, date or style equivalence.",
      "All source parts are retained only in the exact original. No macros, binary formulas, embedded objects or external links execute during inspection.",
      n === "xltx"
        ? "The native template adapter uses the supported XLSX subset. Active or unsupported content may refuse import; review conversion losses."
        : "An external source application is required for editable import and target-format export. No external conversion has been performed here.",
      ...(n === "xls"
        ? [
            "CFB inventory lists directory names; non-workbook stream sizes, nested paths and cell contents are not resolved.",
          ]
        : []),
    ],
  };
  return (DN(a, e, t), a);
}
var UN = Bi(Mo(), 1);
function FN(e, t) {
  let n = (a, s) => {
      if (!e[a]) return;
      let l = new UN.SaxesParser({ xmlns: true });
      (l.on("opentag", s),
        l.on("doctype", () => {
          throw Error("External link metadata cannot contain DTDs.");
        }),
        l.write(Ao(e[a])).close());
    },
    i = new Map();
  n("xl/_rels/workbook.xml.rels", (a) => {
    a.local === "Relationship" &&
      a.attributes.Type?.value.endsWith("/externalLink") &&
      a.attributes.TargetMode?.value !== "External" &&
      i.set(
        a.attributes.Id.value,
        Zs("xl/_rels/workbook.xml.rels", a.attributes.Target.value),
      );
  });
  let r = [];
  if (
    (n("xl/workbook.xml", (a) => {
      if (a.local === "externalReference") {
        let s = Object.values(a.attributes).find(
          (l) => l.local === "id",
        )?.value;
        s && r.push(s);
      }
    }),
    r.length > 100)
  )
    throw Error(
      "More than 100 external workbook targets require an explicit source conversion.",
    );
  let o = [];
  return (
    r.forEach((a, s) => {
      let l = i.get(a);
      if (!l) return;
      let c = l.lastIndexOf("/"),
        u = l.slice(0, c + 1) + "_rels/" + l.slice(c + 1) + ".rels",
        d;
      (n(u, (f) => {
        if (
          f.local === "Relationship" &&
          f.attributes.Type?.value.endsWith("/externalLinkPath") &&
          f.attributes.TargetMode?.value === "External"
        ) {
          if (d !== undefined)
            throw Error("Ambiguous external workbook target relationship.");
          d = f.attributes.Target?.value;
        }
      }),
        d &&
          o.push({ key: `[${s + 1}]`, target: d, part: l, sourceSha256: t }));
    }),
    o
  );
}
function rq(e, t, n) {
  let i = vk(e);
  if (xg(i) !== t)
    throw Error(
      "The workbook content does not match the template conversion route.",
    );
  if (
    ((i["[Content_Types].xml"] = za(
      Ao(i["[Content_Types].xml"]).replace(wg[t], wg[n]),
    )),
    xg(i) !== n)
  )
    throw Error("The template content type could not be converted safely.");
  return Rd(i);
}
async function zN(e) {
  if (e.size > 10 * 1024 * 1024) throw Error("Files must be 10 MB or smaller.");
  let t = new Uint8Array(await e.arrayBuffer()),
    n = Vd(t, e.name);
  if (n.format !== "xltx")
    throw Error("Only XLTX templates use the native template adapter.");
  let i = rq(t, "xltx", "xlsx"),
    r = e.name.replace(/\.xltx$/i, "").slice(0, 236) + ".template-copy.xlsx",
    o = await kg({
      name: r,
      size: i.length,
      arrayBuffer: async () => i.buffer,
    });
  ((o.workbook.name =
    e.name.replace(/\.xltx$/i, "").slice(0, 200) || "Template workbook"),
    (o.workbook.alternateSource = dk(t, e.name, "xltx", {
      method: "template-adapter",
      convertedFileName: r,
      convertedSha256: rr(i),
      reviewedAt: new Date().toISOString(),
      equivalence: "not-verified",
    })));
  let a =
    "A new workbook was created from the XLTX template using the supported XLSX subset. The exact template is retained under File > Workbook formats. The XLSX part inventory refers to the converted intermediate, not the original template. Template fidelity requires comparison in the source application.";
  return (
    (o.warnings = [...new Set([...o.warnings, a])]),
    (o.report.warnings = o.warnings),
    (o.report.alternateOriginal = {
      fileName: n.fileName,
      byteLength: n.fileBytes,
      sha256: n.sha256,
      format: n.format,
    }),
    (o.report.fileName = e.name),
    (o.report.fileBytes = t.length),
    (o.report.format = "xltx"),
    (o.report.workbookName = o.workbook.name),
    o
  );
}
var jN = Bi(Mo(), 1);
var _g = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
var nq = "application/vnd.openxmlformats-officedocument.custom-properties+xml",
  wk = "application/vnd.openxmlformats-officedocument.customXmlProperties+xml",
  iq = { scope: "workbook" };
function Ic(e, t, n, i) {
  let r = new jN.SaxesParser({ xmlns: true });
  (r.on("opentag", t),
    n && r.on("closetag", n),
    i && r.on("text", i),
    r.on("doctype", () => {
      throw Error("DTD is not permitted in preservation metadata.");
    }),
    r.write(Ao(e)).close());
}
function WN(e) {
  return e === "_rels/.rels"
    ? ""
    : e.replace(/(^|\/)\_rels\/([^/]+)\.rels$/, "$1$2");
}
function BN(e) {
  if (!e) return "_rels/.rels";
  let t = e.lastIndexOf("/");
  return `${e.slice(0, t + 1)}_rels/${e.slice(t + 1)}.rels`;
}
function oq(e) {
  let t = new Map(),
    n = new Map(),
    i = [],
    r = new Map();
  Ic(e["[Content_Types].xml"], (o) => {
    (o.local === "Override" &&
      t.set(
        (o.attributes.PartName?.value ?? "").slice(1),
        o.attributes.ContentType?.value ?? "",
      ),
      o.local === "Default" &&
        n.set(
          o.attributes.Extension?.value ?? "",
          o.attributes.ContentType?.value ?? "",
        ));
  });
  for (let [o, a] of Object.entries(e))
    o.endsWith(".rels") &&
      Ic(a, (s) => {
        if (s.local !== "Relationship") return;
        let l = s.attributes.TargetMode?.value === "External",
          c = s.attributes.Target?.value ?? "";
        i.push({
          owner: WN(o),
          rels: o,
          id: s.attributes.Id?.value ?? "",
          type: s.attributes.Type?.value ?? "",
          external: l,
          target: l ? c : Zs(o, c),
        });
      });
  return (
    Ic(e["xl/workbook.xml"], (o) => {
      if (o.local === "sheet") {
        let a = i.find(
          (s) =>
            s.owner === "xl/workbook.xml" &&
            s.id ===
              Object.values(o.attributes).find(
                (l) => l.local === "id" && l.uri === _g.slice(0, -1),
              )?.value,
        );
        a && r.set(a.target, o.attributes.name?.value ?? "Unknown worksheet");
      }
    }),
    {
      relationships: i,
      sheets: r,
      contentType: (o) =>
        t.get(o) ?? n.get(o.split(".").at(-1)) ?? "undeclared",
      types: t,
    }
  );
}
function sq(e, t) {
  return e.endsWith("/")
    ? "directory"
    : e === "[Content_Types].xml" || e.endsWith(".rels")
      ? "package relationships"
      : /vba|macrosheet|activeX|embedding/i.test(e + t)
        ? "active content"
        : /^customXml\//.test(e)
          ? "custom XML"
          : e === "docProps/custom.xml"
            ? "custom properties"
            : /^xl\/charts\//.test(e)
              ? "charts"
              : /^xl\/(?:model|dataModel)\//i.test(e) ||
                  /model\+data|dataModel/i.test(t)
                ? "Excel Data Model"
                : /^xl\/(pivotTables|pivotCache)\//.test(e)
                  ? "pivots"
                  : /connections|queryTables|externalLinks/.test(e)
                    ? "connections and external links"
                    : /slicers|slicerCaches|timelines/.test(e)
                      ? "slicers and timelines"
                      : /threadedComments|persons/.test(e)
                        ? "threaded comments"
                        : /^_xmlsignatures\//.test(e)
                          ? "signatures"
                          : /^xl\/worksheets\/[^/]+\.xml$/.test(e)
                            ? "worksheet"
                            : e === "xl/workbook.xml"
                              ? "workbook and names"
                              : /^xl\/tables\//.test(e)
                                ? "tables"
                                : /^xl\/drawings\//.test(e)
                                  ? "drawings"
                                  : /^xl\/media\//.test(e)
                                    ? "media"
                                    : /^xl\/comments/.test(e)
                                      ? "notes"
                                      : e === "xl/sharedStrings.xml"
                                        ? "shared strings"
                                        : e === "xl/styles.xml" ||
                                            /^xl\/theme\//.test(e)
                                          ? "formatting"
                                          : /^docProps\//.test(e)
                                            ? "document properties"
                                            : "unknown part";
}
function xk(e, t) {
  let n = oq(e),
    i = new Map(),
    r = new Map();
  for (let p of n.relationships)
    (i.set(p.owner, [...(i.get(p.owner) ?? []), p]),
      p.external || r.set(p.target, [...(r.get(p.target) ?? []), p]));
  let o = new Map();
  for (let [p, h] of n.sheets) {
    let m = [p],
      g = new Set();
    for (; m.length;) {
      let b = m.pop();
      if (g.has(b)) continue;
      g.add(b);
      let w = o.get(b) ?? new Set();
      (w.add(h), o.set(b, w));
      for (let _ of i.get(b) ?? []) _.external || m.push(_.target);
    }
  }
  let a = new Set(),
    s = new Map(),
    l = Object.keys(e).filter(
      (p) => p.startsWith("customXml/") && !p.endsWith("/"),
    );
  if (l.length) {
    let p = new Set(l),
      h = n.relationships.filter(
        (_) => p.has(_.owner) || (!_.external && p.has(_.target)),
      ),
      m = !Object.keys(e).some((_) => /xmlMaps|singleXmlCells/i.test(_));
    ((m &&= l.every(
      (_) =>
        _.endsWith(".rels") ||
        (/\.xml$/.test(_) &&
          ["application/xml", "text/xml", wk].includes(n.contentType(_))),
    )),
      (m &&= h.every(
        (_) =>
          !_.external &&
          ((_.owner === "xl/workbook.xml" &&
            _.type === _g + "customXml" &&
            p.has(_.target)) ||
            (p.has(_.owner) &&
              p.has(_.target) &&
              _.type === _g + "customXmlProps")),
      )),
      (m &&= h.every((_) =>
        _.owner === "xl/workbook.xml"
          ? ["application/xml", "text/xml"].includes(n.contentType(_.target))
          : n.contentType(_.owner) !== wk && n.contentType(_.target) === wk,
      )));
    let g = h.filter((_) => _.owner === "xl/workbook.xml"),
      b = new Set(),
      w = g.map((_) => _.target);
    for (; w.length;) {
      let _ = w.pop();
      if (!b.has(_)) {
        (b.add(_), e[BN(_)] && b.add(BN(_)));
        for (let S of i.get(_) ?? []) S.external || w.push(S.target);
      }
    }
    m &&= g.length > 0 && l.every((_) => b.has(_));
    for (let _ of l)
      m
        ? a.add(_)
        : s.set(
            _,
            "Custom XML has bindings, unrecognized types, orphaned parts or relationships outside the passive metadata island. Retained only in the original.",
          );
  }
  let c = "docProps/custom.xml";
  if (e[c]) {
    let p = r.get(c) ?? [],
      h =
        n.contentType(c) === nq &&
        p.length === 1 &&
        p[0].owner === "" &&
        p[0].type === _g + "custom-properties" &&
        !i.get(c)?.length;
    (Ic(e[c], (m) => {
      (m.local === "property" && m.attributes.linkTarget && (h = false),
        (m.uri ===
          "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" &&
          ["Properties", "property"].includes(m.local)) ||
          (m.uri ===
            "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes" &&
            [
              "lpwstr",
              "lpstr",
              "bstr",
              "i1",
              "i2",
              "i4",
              "i8",
              "int",
              "ui1",
              "ui2",
              "ui4",
              "ui8",
              "uint",
              "r4",
              "r8",
              "decimal",
              "bool",
              "date",
              "filetime",
              "empty",
              "null",
            ].includes(m.local)) ||
          (h = false));
    }),
      h
        ? a.add(c)
        : s.set(
            c,
            "Custom properties contain links, unsupported value types or relationships; retained only in the original.",
          ));
  }
  let u = new Set();
  if (e["xl/sharedStrings.xml"]) {
    let p = -1;
    Ic(e["xl/sharedStrings.xml"], (h) => {
      (h.local === "si" && p++, h.local === "r" && u.add(p));
    });
  }
  let d = 0,
    f = Object.entries(e)
      .sort(([p], [h]) => p.localeCompare(h))
      .map(([p, h]) => {
        let m = sq(p, n.contentType(p)),
          g = [...(o.get(p.endsWith(".rels") ? WN(p) : p) ?? [])].map((O) => ({
            sheet: O,
            scope: "sheet",
          }));
        g.length ||
          g.push(
            [
              "charts",
              "pivots",
              "drawings",
              "notes",
              "unknown part",
              "media",
            ].includes(m)
              ? {
                  scope: "unresolved",
                  detail:
                    "No worksheet relationship identifies the affected cells.",
                }
              : iq,
          );
        let b = "dropped",
          w =
            "No safe editable mapping. Exact source bytes remain in the original file.";
        (m === "Excel Data Model" &&
          (w =
            "Imported Excel Data Model, relationships and DAX are not executed or translated. Worksheet results are historical snapshots. Exact model bytes remain only in the original source; rebuild reviewed local typed tables and relationships to calculate new results."),
          ["worksheet", "workbook and names", "tables", "notes"].includes(m) &&
            ((b = "editable"),
            (w =
              "Supported content is editable and regenerated; inspect the feature details and conversion notices for losses. XML bytes are not preserved.")),
          [
            "package relationships",
            "shared strings",
            "formatting",
            "document properties",
            "drawings",
          ].includes(m) &&
            ((b = "flattened"),
            (w =
              "Supported content is regenerated or flattened. Unsupported settings and original XML are retained only in the source.")),
          m === "media" &&
            /^image\/(png|jpeg)$/.test(n.contentType(p)) &&
            ((b = "flattened"),
            (w =
              "Supported worksheet pictures are re-encoded with their supported placement; orphaned images and chart assets are not carried over.")),
          a.has(p) &&
            ((b = "preserved-opaque"),
            (w =
              "Passive metadata snapshot copied byte-for-byte, without interpreting or refreshing its contents. Its original meaning may be stale after edits.")),
          s.has(p) && (w = s.get(p)),
          m === "signatures" &&
            (w =
              "Signatures cannot authenticate an edited workbook; retained only in the original file."));
        let _ = [],
          S = (O, E, R, I, U) => {
            if (++d > sg.maxDetails)
              throw Error(
                "XLSX compatibility details exceed the 100,000-entry budget. Split the source workbook before importing.",
              );
            _.push({
              feature: O,
              disposition: E,
              location: {
                ...g[0],
                ...(R ? { range: R, scope: "cell" } : {}),
                ...(I ? { detail: I } : {}),
              },
              reason:
                U ??
                (E === "editable"
                  ? "Supported subset is editable; review conversion notices for unsupported attributes."
                  : "This source feature is not preserved as an editable feature in the derived workbook."),
            });
          };
        if (
          m === "worksheet" ||
          m === "workbook and names" ||
          m === "tables" ||
          m === "drawings" ||
          m === "pivots" ||
          m === "charts"
        ) {
          let O = "",
            E = "",
            R,
            I = "",
            U = false,
            P = false,
            T = "";
          Ic(
            h,
            (A) => {
              let L = (y) => A.attributes[y]?.value;
              (A.local === "c" &&
                ((O = L("r") ?? ""),
                (E = L("t") ?? ""),
                E === "e" && S("error cell", "flattened", O)),
                A.local === "v" && ((P = true), (T = "")),
                A.local === "f" && ((U = true), (I = "")),
                A.local === "definedName" &&
                  ((R = L("name")),
                  S("defined name", "editable", undefined, R)),
                A.local === "dataValidation" &&
                  S("data validation", "editable", L("sqref"), L("type")),
                A.local === "conditionalFormatting" &&
                  S("conditional formatting", "editable", L("sqref")),
                A.local === "mergeCell" &&
                  S("merged cells", "editable", L("ref")),
                A.local === "table" &&
                  S("structured table", "editable", L("ref"), L("name")),
                A.local === "location" &&
                  S("pivot output", "dropped", L("ref")),
                A.local === "worksheetSource" &&
                  S("pivot source", "dropped", L("ref"), L("sheet")),
                A.local === "r" &&
                  O &&
                  E === "inlineStr" &&
                  S("rich text", "flattened", O),
                [
                  "sheetProtection",
                  "workbookProtection",
                  "autoFilter",
                  "extLst",
                  "hyperlink",
                  "legacyDrawingHF",
                  "webPublishItems",
                  "controls",
                  "oleObjects",
                  "pivotTable",
                  "pivotTablePart",
                  "slicer",
                  "timeline",
                  "sp",
                  "graphicFrame",
                ].includes(A.local) &&
                  S(A.local, "dropped", L("ref"), L("name")));
            },
            (A) => {
              (A.local === "v" &&
                ((P = false),
                O &&
                  E === "s" &&
                  u.has(Number(T)) &&
                  S("rich text", "flattened", O)),
                A.local === "c" && (O = ""),
                A.local === "f" &&
                  ((U = false),
                  m === "charts"
                    ? S(
                        "chart source reference",
                        "dropped",
                        undefined,
                        I.slice(0, 1024),
                      )
                    : /(?:\[|https?:|file:|WEBSERVICE|HYPERLINK|RTD|DDE|CALL\s*\(|EXEC\s*\()/i.test(
                        I,
                      ) &&
                      S(
                        "external/action formula",
                        "flattened",
                        O,
                        undefined,
                        "External/action references may be neutralized to literal text; the formula cache report provides the actual conversion outcome.",
                      )));
            },
            (A) => {
              (U && (I += A), P && (T += A));
            },
          );
        }
        for (let O of i.get(p) ?? [])
          O.external &&
            S(
              "external relationship",
              "dropped",
              undefined,
              O.type.split("/").at(-1),
              "External target is never fetched or carried into the derived copy; its URI remains only in the original.",
            );
        return {
          path: p,
          contentType: p.endsWith("/") ? "directory" : n.contentType(p),
          bytes: h.length,
          sha256: rr(h),
          family: m,
          disposition: b,
          reason: w,
          locations: g,
          features: _,
          retainedInOriginal: true,
        };
      });
  return aq({
    schemaVersion: 1,
    policy: "passive-metadata-v1",
    complete: true,
    phase: "eligible-for-export",
    sourceSha256: t,
    parts: f,
    counts: { editable: 0, "preserved-opaque": 0, flattened: 0, dropped: 0 },
    opaqueBytes: 0,
    note: "All package entries are inventoried. Locations come from source relationships and XML; unresolved locations are explicit. Dispositions describe the derived copy, while all parts remain byte-exact in the retained original. Editable parts can also contain flattened or dropped features. No source part is executed or fetched.",
  });
}
function aq(e) {
  ((e.counts = {
    editable: 0,
    "preserved-opaque": 0,
    flattened: 0,
    dropped: 0,
  }),
    (e.opaqueBytes = 0));
  for (let t of e.parts)
    (e.counts[t.disposition]++,
      t.disposition === "preserved-opaque" && (e.opaqueBytes += t.bytes));
  return e;
}
var zie = Object.freeze({
  storage: "IEEE-754 binary64",
  excelSignificantDigits: 15,
  rounding:
    "ROUND uses decimal ties away from zero; monetary rounding is explicit, never automatic.",
  display:
    "General displays 15 significant digits. Display precision does not round stored values or calculations.",
});
function VN(e) {
  let t = /^[+-]?(\d+\.?\d*|\.\d+)(?:e([+-]?\d+))?$/i.exec(e.trim());
  if (!t) return;
  let n = t[1].replace(".", "").replace(/^0+/, "").replace(/0+$/, ""),
    i = Number(e);
  if (!Number.isFinite(i)) return "outside the finite numeric range";
  if (n && (i === 0 || Math.abs(i) < 22250738585072014e-324))
    return "below Excel's normal floating-point range";
  if (n.length > 15) return "more than 15 significant digits";
  if (Number.isInteger(i) && !Number.isSafeInteger(i))
    return "outside the exact integer range";
}
var qN = Bi(Mo(), 1);
function XN(e) {
  let t = new Map(),
    n = new qN.SaxesParser({ xmlns: true }),
    i = "",
    r = "",
    o = "",
    a = false;
  n.on("opentag", (l) => {
    l.local === "c"
      ? ((i = String(l.attributes.r?.value ?? "")),
        (r = String(l.attributes.t?.value ?? "n")))
      : l.local === "v" && i && ((a = true), (o = ""));
  });
  let s = (l) => {
    if (a && ((o += l), o.length > 1024))
      throw Error("An XLSX numeric token exceeds 1,024 characters.");
  };
  return (
    n.on("text", s),
    n.on("cdata", s),
    n.on("closetag", (l) => {
      if (l.local === "v") {
        if (r === "n" && o.trim()) {
          let c = o.trim();
          if (
            !/^[+-]?(?:\d+\.?\d*|\.\d+)(?:e[+-]?\d+)?$/i.test(c) ||
            !Number.isFinite(Number(c))
          )
            throw Error(`Invalid finite numeric value at ${i}.`);
          t.set(i, c);
        }
        a = false;
      } else l.local === "c" && (i = "");
    }),
    n.write(e).close(),
    t
  );
}
var GN = () => crypto.randomUUID(),
  Ks = (e, t) => `${Ue(t)}${e + 1}`,
  Xd = (e) => {
    let t = Qe(e);
    return `${Ks(t.firstRow, t.firstCol)}:${Ks(t.lastRow, t.lastCol)}`;
  };
var qd = (e, t) =>
  e?.argb && /^[a-f\d]{8}$/i.test(e.argb)
    ? `#${e.argb.slice(-6).toLowerCase()}`
    : t;
function lq(e) {
  let t = e.map(tt).sort((o, a) => o.row - a.row || o.col - a.col),
    n = [];
  for (let o of t) {
    let a = n.at(-1);
    a && a.lastRow === o.row && a.lastCol + 1 === o.col
      ? (a.lastCol = o.col)
      : n.push({
          firstRow: o.row,
          lastRow: o.row,
          firstCol: o.col,
          lastCol: o.col,
        });
  }
  let i = [],
    r = new Map();
  for (let o of n) {
    let a = `${o.firstCol}:${o.lastCol}`,
      s = r.get(a);
    s && s.lastRow + 1 === o.firstRow
      ? (s.lastRow = o.lastRow)
      : (i.push(o), r.set(a, o));
  }
  return i.map(
    (o) => `${Ks(o.firstRow, o.firstCol)}:${Ks(o.lastRow, o.lastCol)}`,
  );
}
function ZN(e) {
  let t = [];
  for (let n of e.matchAll(
    /<(?:\w+:)?dataValidation\b([^>]*?)(?:\/>|>([\s\S]*?)<\/(?:\w+:)?dataValidation>)/g,
  )) {
    let i = (s) =>
        new RegExp("\\b" + s + `=["\\']([^"\\']*)["\\']`).exec(n[1])?.[1],
      r = (i("sqref") ?? "").trim().split(/\s+/).filter(Boolean);
    if (!r.length) continue;
    let o = Xd(r[0]).split(":")[0];
    if (t.length + r.length > 1e3)
      throw Error(
        "XLSX has more than 1,000 validation areas; narrow the source rules before import.",
      );
    let a = [
      ...(n[2] ?? "").matchAll(
        /<(?:\w+:)?formula[12]\b[^>]*>([\s\S]*?)<\/(?:\w+:)?formula[12]>/g,
      ),
    ].map((s) => s[1].trim());
    for (let s of r)
      t.push({
        range: Xd(s),
        anchor: o,
        operands: a,
        dropdown: !["1", "true"].includes(i("showDropDown") ?? "0"),
        alert: ["1", "true"].includes(i("showErrorMessage") ?? "0"),
      });
  }
  return t;
}
var cq = {
  between: "between",
  notBetween: "notBetween",
  greater: "greaterThan",
  greaterEqual: "greaterThanOrEqual",
  less: "lessThan",
  lessEqual: "lessThanOrEqual",
  equal: "equal",
  notEqual: "notEqual",
};
function uq(e) {
  let t = {
      list: "list",
      decimal: "number",
      whole: "integer",
      textLength: "textLength",
      date: "date",
      time: "time",
      custom: "custom",
    }[e.type],
    n = Object.entries(cq).find(([, u]) => u === e.operator)?.[0];
  if (!t || (!n && !["list", "custom"].includes(t))) return null;
  let i = e.formulae ?? [],
    r = (u) =>
      u instanceof Date
        ? (u.getTime() - Date.UTC(1899, 11, 30)) / 864e5
        : typeof u == "number"
          ? u
          : t === "time" &&
              typeof u == "string" &&
              /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?$/i.test(u)
            ? Number(u)
            : NaN,
    o = ["list", "custom"].includes(t) ? 0 : r(i[0]),
    a = ["list", "custom"].includes(t)
      ? 0
      : ["between", "notBetween"].includes(n)
        ? r(i[1])
        : o;
  if (!Number.isFinite(o) || !Number.isFinite(a)) return null;
  let s = [],
    l;
  if (
    (t === "list" &&
      (typeof i[0] != "string" ||
        !/^"[^"]*"$/.test(i[0]) ||
        ((s = i[0].slice(1, -1).split(",")),
        !s.length ||
          s.length > 200 ||
          s.some((u) => !u.length || u.length > 1e3)))) ||
    (t === "custom" &&
      (typeof i[0] != "string" ||
        ((l = "=" + i[0].replace(/^=/, "")),
        l.length > 8192 || l.length < 2 || kk(l))))
  )
    return null;
  let c = e.errorStyle?.toLowerCase() ?? "stop";
  return ["stop", "warning", "information"].includes(c)
    ? {
        kind: t,
        operator: n ?? "equal",
        min: o,
        max: a,
        list: s,
        allowBlank: !!e.allowBlank,
        message: String(e.error ?? "Choose a valid value.").slice(0, 500),
        errorStyle: c,
        errorTitle: String(e.errorTitle ?? "Invalid value").slice(0, 100),
        showInputMessage: !!e.showInputMessage,
        inputTitle: String(e.promptTitle ?? "").slice(0, 100),
        inputMessage: String(e.prompt ?? "").slice(0, 500),
        ...(l ? { formula: l } : {}),
      }
    : null;
}
var kk = (e) => {
    let t = Co(e).replace(/"(?:[^"]|"")*"/g, '""');
    return (
      /[\[\]|]/.test(t) ||
      /(?:https?:|file:|javascript:|data:|\\\\)/i.test(t) ||
      /\b(?:WEBSERVICE|HYPERLINK|IMAGE|RTD|DDE|CALL|EXEC|REGISTER(?:\.ID)?|RUN|SQL\.REQUEST|STOCKHISTORY)\s*\(/i.test(
        t.replace(/'(?:[^']|'')*'(?=\s*!)/g, "Sheet"),
      )
    );
  },
  KN = (e) => '"' + e.replace(/"/g, '""') + '"';
function dq(e, t) {
  return "ISNUMBER(FIND(LOWER(" + KN(e) + "),LOWER(" + t + '&"")))';
}
function HN(e, t) {
  let n = Qe(t),
    i = Ks(n.firstRow, n.firstCol),
    r = Xd(t).replace(/([A-Z]+)(\d+)/g, "$$$1$$$2"),
    o =
      "SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(" +
      i +
      ',"~","~~"),"*","~*"),"?","~?")';
  return (
    "AND(LEN(" +
    i +
    ")>0,COUNTIF(" +
    r +
    ',"="&' +
    o +
    ")" +
    (e === "duplicate" ? ">1" : "=1") +
    ")"
  );
}
function fq(e) {
  if (e === null || e === "") return true;
  if (typeof e == "number") return Number.isFinite(e);
  if (typeof e != "string") return false;
  let t = e.trim() !== "" && Number.isFinite(Number(e));
  return !(
    /[^\x20-\x7e]/.test(e) ||
    /^#(?:NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|ERROR!|CYCLE!|SPILL!|CALC!|LIMIT!|BLOCKED!)/.test(
      e,
    ) ||
    e.replace(/[~*?]/g, "~~").length > 254 ||
    (t && String(Number(e)) !== e) ||
    (!t && /^[\s+-]*\d/.test(e))
  );
}
function JN(e) {
  let t = {},
    n = new Set();
  for (let i of e.matchAll(/<(?:\w+:)?cfRule\b([^>]*)>/g)) {
    let r = /\bpriority=["']([^"']*)["']/.exec(i[1])?.[1],
      o = /\bstopIfTrue=["']([^"']+)["']/.exec(i[1])?.[1];
    if (r === undefined && /\bid=["']/.test(i[1])) continue;
    let a = Number(r);
    if (!r || !/^\d+$/.test(r) || !Number.isSafeInteger(a) || a < 1 || n.has(a))
      throw Error("Invalid or duplicate conditional-format priority.");
    if ((n.add(a), r && o !== undefined)) {
      if (!["1", "0", "true", "false"].includes(o))
        throw Error("Invalid conditional-format stop flag.");
      t[String(a)] = o === "1" || o === "true";
    }
  }
  for (let i of e.matchAll(/<(?:\w+:)?cfvo\b([^>]*)>/g)) {
    let r = /\bgte=["']([^"']*)["']/.exec(i[1])?.[1];
    if (r !== undefined && r !== "1" && r !== "true")
      throw Error(
        "Strict or malformed conditional icon thresholds are not supported. Use inclusive thresholds or native JSON.",
      );
  }
  return t;
}
function hq(e, t) {
  let n =
      e.style?.fill?.type === "pattern"
        ? qd(e.style.fill.fgColor, "#fff2cc")
        : "#fff2cc",
    i = qd(e.style?.font?.color, "#000000"),
    r = {
      id: GN(),
      priority: Number.isFinite(e.priority) ? e.priority : 1e4,
      range: Xd(t),
      value: "",
      second: "",
      fill: n,
      color: i,
    },
    o = (c) =>
      typeof c == "number" && Number.isFinite(c)
        ? String(c)
        : typeof c == "string" && c.trim() !== "" && Number.isFinite(Number(c))
          ? c
          : typeof c == "string" && /^"(?:[^"]|"")*"$/.test(c)
            ? c.slice(1, -1).replace(/""/g, '"')
            : null;
  if (
    e.type === "containsText" &&
    (!e.operator || e.operator === "containsText")
  ) {
    let c = Qe(t),
      u = Ks(c.firstRow, c.firstCol),
      d = typeof e.text == "string" ? e.text : "",
      f = e.formulae?.[0],
      p =
        typeof f == "string" && f.trim()
          ? f.replace(/^=/, "")
          : "NOT(ISERROR(SEARCH(" + KN(d) + "," + u + ")))";
    return p.length > 4095 || kk("=" + p)
      ? null
      : { ...r, kind: "formula", formula: "=" + p };
  }
  if (
    e.type === "colorScale" &&
    e.cfvo?.length === 2 &&
    e.cfvo[0].type === "min" &&
    e.cfvo[1].type === "max"
  )
    return {
      ...r,
      kind: "colorScale",
      minColor: qd(e.color?.[0], "#f8e99b"),
      maxColor: qd(e.color?.[1], "#63be7b"),
    };
  if (
    e.type === "dataBar" &&
    e.cfvo?.length === 2 &&
    ["min", "autoMin"].includes(e.cfvo[0].type) &&
    ["max", "autoMax"].includes(e.cfvo[1].type) &&
    e.showValue !== false &&
    e.direction !== "rightToLeft"
  )
    return { ...r, kind: "dataBar", fill: qd(e.color, "#4472c4") };
  if (
    e.type === "iconSet" &&
    e.iconSet === "3Arrows" &&
    !e.reverse &&
    !e.custom &&
    e.showValue !== false &&
    e.cfvo?.length === 3 &&
    e.cfvo[1].type === e.cfvo[2].type &&
    ["num", "percent"].includes(e.cfvo[1].type)
  ) {
    let [c, u, d] = e.cfvo,
      f = u.value,
      p = d.value;
    return !(c.type === "min" || (c.type === "percent" && c.value === 0)) ||
      typeof f != "number" ||
      typeof p != "number" ||
      !Number.isFinite(f) ||
      !Number.isFinite(p) ||
      f >= p ||
      (u.type === "percent" && (f < 0 || p > 100))
      ? null
      : {
          ...r,
          kind: "iconSet",
          thresholdType: u.type === "num" ? "number" : "percent",
          iconThresholds: [f, p],
        };
  }
  if (e.type === "expression") {
    let c = String(e.formulae?.[0] ?? "").replace(/^=/, "");
    if (!c || c.length > 4095 || kk("=" + c)) return null;
    if (c.toUpperCase() === HN("duplicate", t).toUpperCase())
      return { ...r, kind: "duplicate" };
    if (c.toUpperCase() === HN("unique", t).toUpperCase())
      return { ...r, kind: "unique" };
    let u =
      /^ISNUMBER\(FIND\(LOWER\("((?:[^"]|"")*)"\),LOWER\(([A-Z]+[1-9]\d*)&""\)\)\)$/i.exec(
        c,
      );
    if (u) {
      let h = u[1].replace(/""/g, '"'),
        m = Qe(t),
        g = Ks(m.firstRow, m.firstCol);
      if (h.length <= 1e3 && c.toUpperCase() === dq(h, g).toUpperCase())
        return { ...r, kind: "contains", value: h };
    }
    let d = Qe(t),
      f = Ks(d.firstRow, d.firstCol),
      p = {
        [`LEN(${f})=0`]: "blank",
        [`LEN(${f})>0`]: "notBlank",
        [`ISERROR(${f})`]: "error",
        [`NOT(ISERROR(${f}))`]: "notError",
      };
    return p[c]
      ? { ...r, kind: p[c] }
      : { ...r, kind: "formula", formula: "=" + c };
  }
  if (e.type !== "cellIs") return null;
  let a = {
      greaterThan: "greater",
      lessThan: "less",
      equal: "equal",
      between: "between",
    }[e.operator],
    s = o(e.formulae?.[0]),
    l = a === "between" ? o(e.formulae?.[1]) : "";
  return !a || s === null || l === null
    ? null
    : { ...r, kind: a, value: s, second: l };
}
function YN(e, t, n = {}, i = []) {
  let r = [],
    o = {},
    a = e.model.merges ?? [];
  a.length && (o.merges = a.map(Xd));
  let s = e.dataValidations.model,
    l = new Map();
  for (let [h, m] of Object.entries(s)) {
    let g = tt(h),
      b = i.find((_) => {
        let S = Qe(_.range);
        return (
          g.row >= S.firstRow &&
          g.row <= S.lastRow &&
          g.col >= S.firstCol &&
          g.col <= S.lastCol
        );
      }),
      w = JSON.stringify({ rule: m, area: b });
    (l.has(w) || l.set(w, { rule: m, addresses: [], area: b }),
      l.get(w).addresses.push(h));
  }
  let c = [];
  for (let { rule, addresses, area } of l.values()) {
    let b = !["list", "custom"].includes(rule.type),
      w = area?.operands.every((S) =>
        /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?$/i.test(S),
      ),
      _ =
        b && area && !w
          ? null
          : uq(
              b && area
                ? { ...rule, formulae: area.operands.map(Number) }
                : rule,
            );
    if (!_) {
      r.push(
        "An unsupported validation formula, operator or external list was omitted; source cells were retained.",
      );
      continue;
    }
    area &&
      !area.alert &&
      r.push(
        "An imported validation has error alerts disabled. Local validation requires Stop or explicit Warning/Information review; source values were retained.",
      );
    for (let S of lq(addresses)) {
      if (c.length >= 100) {
        r.push(
          "Additional validation areas beyond the 100-rule limit were omitted; source cells were retained.",
        );
        break;
      }
      let O = tt(area?.anchor ?? S.split(":")[0]),
        E = tt(S.split(":")[0]);
      c.push({
        ..._,
        id: GN(),
        range: S,
        ...(area ? { showDropdown: area.dropdown } : {}),
        ...(_.formula
          ? { formula: bi(_.formula, E.row - O.row, E.col - O.col) }
          : {}),
      });
    }
  }
  c.length && (o.validations = c);
  let u = [],
    d = new Set(),
    f = 0;
  for (let h of e.conditionalFormattings ?? []) {
    let m = h.ref.trim().split(/\s+/).filter(Boolean);
    if (!m.length)
      throw Error("A conditional formatting rule requires an applied range.");
    let g = Qe(m[0]);
    for (let b of h.rules) {
      if (
        !Number.isSafeInteger(b.priority) ||
        b.priority < 1 ||
        d.has(b.priority)
      )
        throw Error("Invalid or duplicate conditional-format priority.");
      d.add(b.priority);
      for (let w of m) {
        let _ = Qe(w),
          S = b;
        (b.type === "expression" || b.type === "containsText") &&
          b.formulae?.length &&
          (S = {
            ...b,
            formulae: b.formulae.map((E) =>
              typeof E == "string"
                ? bi(
                    "=" + E.replace(/^=/, ""),
                    _.firstRow - g.firstRow,
                    _.firstCol - g.firstCol,
                  ).slice(1)
                : E,
            ),
          });
        let O = hq(S, w);
        if (O && (O.kind === "duplicate" || O.kind === "unique")) {
          let E = true;
          for (let [R, I] of Object.entries(t.cells)) {
            if (++f > 1e6) {
              E = false;
              break;
            }
            let U = tt(R);
            if (
              U.row < _.firstRow ||
              U.row > _.lastRow ||
              U.col < _.firstCol ||
              U.col > _.lastCol
            )
              continue;
            let P = I.raw;
            if (
              (typeof I.raw == "string" &&
                I.raw.startsWith("=") &&
                !I.literal &&
                ((P = e.getCell(R).result),
                P === undefined && /^=\s*""\s*$/.test(I.raw) && (P = "")),
              !fq(P))
            ) {
              E = false;
              break;
            }
          }
          E ||
            ((O = {
              ...O,
              kind: "formula",
              formula: "=" + String(S.formulae?.[0] ?? "").replace(/^=/, ""),
            }),
            r.push(
              "A duplicate/unique condition was retained as a formula because its population cannot use the native literal comparison; formula-engine limits apply.",
            ));
        }
        O && u.length < 100
          ? u.push({
              ...O,
              ...(n[String(b.priority)] !== undefined
                ? { stopIfTrue: n[String(b.priority)] }
                : {}),
            })
          : r.push(
              "An unsupported conditional-format rule or a rule beyond the 100-rule limit was omitted; base cell formatting was retained.",
            );
      }
    }
  }
  if (u.length) {
    let h = [...u].sort((g, b) => (g.priority ?? 1e4) - (b.priority ?? 1e4)),
      m = new Map(h.map((g, b) => [g.id, b + 1]));
    o.conditionalFormats = u.map((g) => ({ ...g, priority: m.get(g.id) }));
  }
  return (
    e.model.sheetProtection?.sheet &&
      r.push(
        "Source Excel sheet-password protection is not imported. The local protection feature and host authorization are separate controls.",
      ),
    Rp(o, t),
    r
  );
}
var QN = Bi(Mo(), 1);
function e8(e) {
  let t = {
      mode: "auto",
      iterative: false,
      maxIterations: null,
      maxChange: null,
    },
    n = new QN.SaxesParser({ xmlns: true }),
    i = false;
  return (
    n.on("opentag", (r) => {
      if (r.local !== "calcPr" || !/spreadsheetml/.test(r.uri)) return;
      if (i)
        throw Error(
          "Duplicate XLSX calculation properties. Repair the source workbook.",
        );
      i = true;
      let o = (d) => {
          let f = r.attributes[d];
          return f && typeof f != "string" ? f.value : f;
        },
        a = o("calcMode") ?? "auto",
        s = o("iterate"),
        l = o("iterateCount"),
        c = o("iterateDelta"),
        u = o("fullPrecision");
      if (
        !["auto", "autoNoTable", "manual"].includes(a) ||
        (u !== undefined && !["true", "false", "1", "0"].includes(u)) ||
        (s !== undefined && !["true", "false", "1", "0"].includes(s)) ||
        (l !== undefined &&
          (!/^\d+$/.test(l) ||
            !Number.isSafeInteger(Number(l)) ||
            Number(l) < 1)) ||
        (c !== undefined &&
          (!c.trim() || !Number.isFinite(Number(c)) || Number(c) < 0))
      )
        throw Error(
          "Invalid XLSX calculation properties. Repair the source workbook.",
        );
      t = {
        mode: a,
        iterative: s === "true" || s === "1",
        maxIterations: l === undefined ? null : Number(l),
        maxChange: c === undefined ? null : Number(c),
        ...(u === undefined
          ? {}
          : { fullPrecision: u === "true" || u === "1" }),
      };
    }),
    n.on("doctype", () => {
      throw Error("XLSX calculation properties cannot contain a DTD.");
    }),
    n.write(e).close(),
    t
  );
}
function t8(e, t, n) {
  ((e.calculation = {
    ...js,
    mode: t.mode === "manual" ? "manual" : "automatic",
    iterative: false,
    maxIterations:
      t.maxIterations !== null && t.maxIterations <= 1e3
        ? t.maxIterations
        : js.maxIterations,
    maxChange:
      t.maxChange !== null && t.maxChange <= 1 ? t.maxChange : js.maxChange,
  }),
    t.fullPrecision === false &&
      n.push(
        "The source used precision as displayed. This application calculates with full stored precision and never rounds cells automatically. Source rounding may already have changed stored values irreversibly; review formula results and use explicit ROUND where required. The original setting is retained in imported calculation evidence.",
      ),
    t.mode === "manual" &&
      n.push(
        "Source Manual calculation mode is retained. Formulas remain #PENDING! until Calculate Now; imported caches are historical evidence only.",
      ),
    t.mode === "autoNoTable" &&
      n.push(
        "Excel Automatic except data tables is imported as Automatic. Native Excel data-table calculation is not supported; source mode is retained as evidence.",
      ),
    t.iterative &&
      n.push(
        "Source iterative calculation was enabled. It is disabled locally until you explicitly enable it in Calculation Options and review numeric convergence.",
      ),
    ((t.maxIterations ?? 100) > 1e3 || (t.maxChange ?? 0.001) > 1) &&
      n.push(
        "Source iteration limits exceed local limits. Local defaults are used; original settings remain in imported cache evidence.",
      ));
}
var i8 = Bi(Mo(), 1);
var r8 = new WeakMap();
function n8(e) {
  try {
    let t = new TextDecoder("utf-8", { fatal: true }).decode(e);
    if (t.includes("\0")) throw Error();
    return t;
  } catch {
    throw Error(
      "Automation source must be UTF-8 text without null characters. Export a UTF-8 source copy; binary VBA is inventory-only.",
    );
  }
}
function pq(e) {
  let t = vg(e, { allowBinaryWorkbookInventory: true }),
    n = new Map(),
    i = 0,
    r,
    o = new Set(),
    a = new gc((s) => {
      if (!t.has(s.name) || o.has(s.name)) {
        r = Error("Automation archive directory mismatch.");
        return;
      }
      o.add(s.name);
      let l = [],
        c = 0,
        u = -1;
      ((s.ondata = (d, f, p) => {
        if (!r) {
          if (d) {
            r = Error(`Corrupt automation archive: ${s.name}.`);
            return;
          }
          if (
            ((c += f.length),
            (i += f.length),
            c > t.get(s.name).size || i > 50 * 1024 * 1024)
          ) {
            ((r = Error("Automation archive exceeds its expansion budget.")),
              s.terminate());
            return;
          }
          if (((u = Tc(u, f)), l.push(f), p)) {
            if (
              c !== t.get(s.name).size ||
              (u ^ -1) >>> 0 !== t.get(s.name).crc
            ) {
              r = Error(`Automation archive checksum mismatch: ${s.name}.`);
              return;
            }
            let h = new Uint8Array(c),
              m = 0;
            for (let g of l) (h.set(g, m), (m += g.length));
            n.set(s.name, h);
          }
        }
      }),
        s.start());
    });
  a.register(mc);
  try {
    a.push(e, true);
  } catch {
    throw Error("Unable to inspect corrupt automation archive.");
  }
  if (r) throw r;
  if (n.size !== t.size) throw Error("Automation archive is incomplete.");
  return n;
}
function mq(e, t, n) {
  if (/<\?xml[^>]*encoding\s*=\s*["'](?!utf-8|utf8)[^"']+/i.test(e))
    throw Error(`Unsupported XML encoding in ${t}.`);
  let i = new i8.SaxesParser({ xmlns: false }),
    r = 0,
    o = 0;
  (i.on("doctype", () => {
    throw Error("Automation inventory does not accept XML document types.");
  }),
    i.on("opentag", (a) => {
      if (++r > 128 || ++o > 1e6)
        throw Error("Automation inventory XML exceeds its bounds.");
      n(a.name.split(":").at(-1), a.attributes);
    }),
    i.on("closetag", () => r--),
    i.on("error", (a) => {
      throw Error(`Malformed automation metadata in ${t}: ${a.message}`);
    }),
    i.write(e).close());
}
function o8(e, t) {
  if (
    (cg(t),
    !ArrayBuffer.isView(e) ||
      Object.prototype.toString.call(e) !== "[object Uint8Array]" ||
      !e.length ||
      e.length > jn.maxFileBytes)
  )
    throw Error("Automation files must contain data and be 10 MB or smaller.");
  let n = e.slice(),
    i = {
      schemaVersion: 1,
      fileName: t,
      kind: "excel-container",
      fileBytes: n.length,
      sha256: rr(n),
      retainedEligible: n.length <= jn.maxArtifactBytes,
      execution: "inventory-only",
      items: [],
      warnings: [
        "Imported automation never runs on open, preservation or native JSON reopen. Only a separately reviewed migration can issue bounded workbook commands.",
        "Preserved artifacts are inert original bytes in native JSON. VBA/XLM binaries are not executed or decompiled; retain originals and use an authorized source application to export source modules.",
      ],
    },
    r = t.split(".").at(-1)?.toLowerCase();
  if (["bas", "vba", "cls", "ts", "osts"].includes(r ?? "")) {
    if (n.length > jn.maxSourceBytes)
      throw Error("Text automation source is limited to 256 KiB.");
    let o = n8(n),
      a = r === "ts" || r === "osts" ? "office-script" : "vba";
    if (((i.kind = r === "osts" ? "office-script-container" : a), r === "osts"))
      try {
        let s = JSON.parse(o);
        if (
          !s ||
          typeof s != "object" ||
          Array.isArray(s) ||
          typeof s.body != "string"
        )
          throw Error();
        if (
          ((o = s.body),
          o.includes("\0") ||
            new TextEncoder().encode(o).length > jn.maxSourceBytes)
        )
          throw Error("Source body is outside the supported text budget.");
      } catch {
        (i.warnings.push(
          "Unsupported .osts envelope. Preserve the original or export its code as UTF-8 .ts. This adapter recognizes only a JSON object with a string body; it makes no universal .osts schema claim.",
        ),
          i.items.push({
            path: t,
            kind: "Office Script container",
            bytes: n.length,
            sha256: i.sha256,
            execution: "unsupported",
            note: "Unknown source envelope; no code inferred.",
          }));
      }
    i.items.length ||
      ((i.script = { language: a, text: o }),
      (i.execution = "reviewed-migration-only"),
      i.items.push({
        path: t,
        kind: a === "vba" ? "VBA source" : "Office Script source",
        bytes: n.length,
        sha256: i.sha256,
        execution: "migration-candidate",
        note: "Compatibility scan required. Only literal range operations are eligible; this is not an equivalent VBA or Office Scripts runtime.",
      }));
  } else if (n[0] === 208 && n[1] === 207 && n[2] === 17 && n[3] === 224)
    ((i.kind = "legacy-container"),
      i.items.push({
        path: t,
        kind: "OLE / encrypted or legacy container",
        bytes: n.length,
        sha256: i.sha256,
        execution: "unsupported",
        note: "VBA/XLM presence cannot be determined without a supported decoder. Export a clean OOXML workbook and UTF-8 source modules in an authorized application.",
      }));
  else if (
    ["xlsm", "xlam", "xltm", "xlsx", "xltx"].includes(r ?? "") ||
    (n[0] === 80 && n[1] === 75)
  ) {
    let o = pq(n),
      a = new Map(),
      s = (l, c) => {
        let u = "";
        (/vbaProject/i.test(c) || /vbaproject(?:signature)?\.bin$/i.test(l)
          ? (u = "VBA binary")
          : /(?:intl)?macrosheet/i.test(c) || /(?:^|\/)macrosheets\//i.test(l)
            ? (u = "XLM macro sheet")
            : /office.?script|scriptLink|scriptReference/i.test(c) ||
                /\.(?:osts|ts)$/i.test(l)
              ? (u = "Office Script artifact/reference")
              : (/activeX|oleObject|relationships\/(?:package|control)$/i.test(
                  c,
                ) ||
                  /\/(?:activeX|embeddings)\//i.test(l)) &&
                (u = "ActiveX / embedded object"),
          u &&
            a.set(l, {
              kind: u,
              note:
                u === "Office Script artifact/reference"
                  ? "Package artifact only. Import its separate .ts/.osts source for migration."
                  : "Inventory and opaque preservation only; no binary execution.",
            }));
      };
    for (let [l, c] of o)
      (s(l, ""),
        /\.(?:xml|rels)$/i.test(l) &&
          mq(n8(c), l, (u, d) => {
            if (
              (u === "Override" &&
                d.PartName &&
                s(d.PartName.replace(/^\//, ""), d.ContentType ?? ""),
              u === "Default" && d.Extension)
            )
              for (let f of o.keys())
                f.toLowerCase().endsWith("." + d.Extension.toLowerCase()) &&
                  s(f, d.ContentType ?? "");
            if (
              u === "Relationship" &&
              /vbaProject|(?:intl)?macrosheet|activeX|oleObject|office.?script|scriptLink|scriptReference|relationships\/(?:package|control)$/i.test(
                d.Type ?? "",
              )
            ) {
              if (d.TargetMode === "External") {
                i.items.push({
                  path: `${l} \u2192 ${d.Target ?? "(missing)"}`,
                  kind: "External automation reference",
                  bytes: 0,
                  execution: "external-unresolved",
                  note: "Referenced content was not fetched, authenticated or executed.",
                });
                return;
              }
              let f =
                  l === "_rels/.rels"
                    ? ""
                    : l.slice(0, l.lastIndexOf("/_rels/")) + "/",
                p = [];
              for (let h of ((d.Target ?? "").startsWith("/")
                ? d.Target.slice(1)
                : f + (d.Target ?? "")
              ).split("/"))
                h === ".." ? p.pop() : h && h !== "." && p.push(h);
              s(p.join("/"), d.Type ?? "");
            }
          }));
    for (let [l, c] of a) {
      let u = o.get(l);
      i.items.push({
        path: l,
        kind: c.kind,
        bytes: u?.length ?? 0,
        ...(u ? { sha256: rr(u) } : {}),
        execution: "unsupported",
        note: u
          ? c.note
          : "Referenced part is missing. Inventory is incomplete; repair the original.",
      });
    }
    i.items.length ||
      i.warnings.push(
        "No VBA/XLM or known automation parts were detected in this package. This does not prove the absence of separately stored scripts or automation.",
      );
  } else
    throw Error(
      "Choose VBA .bas/.vba/.cls, Office Script .ts/.osts, or an Excel container for inventory.",
    );
  return (
    i.warnings.push(
      "Office Scripts may be stored separately in OneDrive/SharePoint. Workbook inventory cannot discover those scripts; import the authorized source separately. No linked content is fetched.",
    ),
    i.retainedEligible ||
      i.warnings.push(
        "This file exceeds the 2 MB retained-artifact limit. The inventory is available; preserve the original externally.",
      ),
    r8.set(i, { signature: JSON.stringify(i), bytes: n }),
    i
  );
}
function s8(e) {
  return (e.embeddedObjectInventories ?? []).flatMap((t) =>
    t.items.map((n) => ({
      sourceFile: t.fileName,
      sourceSha256: t.sourceSha256,
      path: n.path,
      kind: n.kind,
      bytes: n.bytes,
      ...(n.sha256 ? { sha256: n.sha256 } : {}),
      disposition: "dropped",
      execution: "never",
      reason:
        "Inventory retained in native JSON only; unsupported OLE/ActiveX payload was not imported and is not emitted in derived XLSX.",
    })),
  );
}
function _k(e, t, n, i, r) {
  return {
    schemaVersion: 1,
    direction: e,
    fileName: t,
    fileBytes: n,
    workbookName: i.name,
    revision: i.revision,
    compatibility: "derived-subset",
    lossless: false,
    originalFileRetained: !!i.xlsxSource,
    ...(i.alternateSource
      ? {
          alternateOriginal: {
            fileName: i.alternateSource.fileName,
            byteLength: i.alternateSource.byteLength,
            sha256: i.alternateSource.sha256,
            format: i.alternateSource.format,
          },
        }
      : {}),
    ...(i.xlsxSource
      ? {
          originalSource: {
            fileName: i.xlsxSource.fileName,
            byteLength: i.xlsxSource.byteLength,
            sha256: i.xlsxSource.sha256,
          },
        }
      : {}),
    sheets: i.sheets.map((o) => ({
      name: o.name,
      storedCells: Object.keys(o.cells).length,
      hidden: !!(o.hidden || o.veryHidden),
    })),
    warnings: [...new Set(r)],
    ...(e === "export" && i.embeddedObjectInventories?.length
      ? { embeddedObjectLosses: s8(i) }
      : {}),
    caches: {
      total: 0,
      pending: 0,
      matched: 0,
      different: 0,
      missing: 0,
      neutralized: 0,
      written: 0,
      omitted: 0,
      issues: [],
      matches: [],
      detailsOmitted: 0,
    },
  };
}
function gq(e, t) {
  if (e != null) {
    if (e instanceof Date) return { type: "number", value: _p(e, t) };
    if (typeof e == "object" && "error" in e && typeof e.error == "string")
      return { type: "error", value: e.error };
    if (typeof e == "string") return { type: "text", value: e };
    if (typeof e == "boolean") return { type: "boolean", value: e };
    if (typeof e == "number" && Number.isFinite(e))
      return { type: "number", value: e };
  }
}
function bq(e, t) {
  return (
    e.type === t.type &&
    (typeof e.value == "number" && typeof t.value == "number"
      ? Math.abs(e.value - t.value) <=
        1e-10 * Math.max(1, Math.abs(e.value), Math.abs(t.value))
      : e.value === t.value)
  );
}
function a8(e, t, n, i = new Map(), r, o = new Map()) {
  let a = yi(t);
  for (let s of t.sheets) {
    let l = e.getWorksheet(s.name);
    for (let [c, u] of Object.entries(s.cells)) {
      let d = l.getCell(c).value;
      if (
        !d ||
        typeof d != "object" ||
        !("formula" in d || "sharedFormula" in d)
      )
        continue;
      let f = i.get(s.name)?.has(c)
          ? { type: "text", value: "" }
          : o.get(s.name)?.has(c)
            ? { type: "number", value: Number(o.get(s.name).get(c)) }
            : gq(l.getCell(c).result, t.datePolicy?.dateSystem ?? "1900"),
        p = a.getTypedValue(s.id, c);
      r &&
        r.entries.push({
          sheetId: s.id,
          address: c,
          formula: String(u.raw),
          ...(f ? { cache: f } : {}),
        });
      let h = u.literal
        ? "neutralized"
        : p.type === "error" && p.value === "#PENDING!"
          ? "pending"
          : f
            ? bq(f, p)
              ? "matched"
              : "different"
            : "missing";
      (n.caches.total++, (n.caches[h] = (n.caches[h] ?? 0) + 1));
      let m = h === "matched" ? n.caches.matches : n.caches.issues;
      m.length < 100
        ? m.push({
            sheet: s.name,
            address: c,
            ...(f ? { source: f } : {}),
            calculated: p,
            status: h,
          })
        : n.caches.detailsOmitted++;
    }
  }
  (n.caches.pending &&
    n.warnings.push(
      `${n.caches.pending} formula cache comparison(s) are pending. Use Calculate Now and download calculation evidence for a current comparison.`,
    ),
    n.caches.different &&
      n.warnings.push(
        `${n.caches.different} source formula cache(s) differ from local recalculation. Review the comparison before relying on results; source caches never override local results.`,
      ),
    n.caches.missing &&
      n.warnings.push(
        `${n.caches.missing} formula(s) have no usable source cache. Local results are not independently verified Excel results.`,
      ),
    n.caches.neutralized &&
      n.warnings.push(
        `${n.caches.neutralized} external/action formula(s) were converted to literal text.`,
      ));
}
function l8(e, t) {
  let n = (r, o) =>
      new RegExp(`(?:^|\\s)${o}\\s*=\\s*(["'])(.*?)\\1`, "i")
        .exec(r)?.[2]
        ?.replace(/&#(x[0-9a-f]+|[0-9]+);/gi, (s, l) =>
          String.fromCodePoint(
            l[0].toLowerCase() === "x" ? parseInt(l.slice(1), 16) : Number(l),
          ),
        ),
    i;
  for (let r of e.matchAll(
    /<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>|<(\/?)(?:\w+:)?(c|f)\b([^>]*)>/gi,
  )) {
    if (!r[2]) continue;
    if (r[2].toLowerCase() === "c") {
      i = r[1] || /\/\s*$/.test(r[3]) ? undefined : r[3];
      continue;
    }
    if (i === undefined || r[1] || n(r[3], "t") !== "array") continue;
    let o = n(i, "r") ?? "unknown address",
      a = n(r[3], "ref") ?? o,
      s =
        n(i, "cm") !== undefined
          ? "Metadata-qualified Excel array formula (dynamic/legacy classification requires Excel metadata)"
          : "Legacy CSE array formula";
    throw new Error(
      `${s} at ${t}!${o}, range ${a}. Import stopped before changing the workbook. Open the original in Excel and convert the array to ordinary supported formulas or an explicitly reviewed values snapshot, then import that copy. Fixed CSE ranges and Excel array metadata are not silently treated as dynamic spills.`,
    );
  }
}
var y8 = Bi(u8(), 1);
var p8 = Bi(d8(), 1);
function Tk(e) {
  throw new Error(e);
}
var yq = 864e5,
  vq = Date.UTC(1899, 11, 31),
  wq = 2958465,
  Ak = class extends Error {};
function xq() {
  throw new Ak("Date series must stay between 1899-12-31 and 9999-12-31.");
}
function kq(e, t, n) {
  let i = new Date(0);
  return (i.setUTCHours(0, 0, 0, 0), i.setUTCFullYear(e, t, n), i);
}
function _q(e) {
  let t = (e.getTime() - vq) / yq,
    n = t >= 60 ? t + 1 : t;
  return ((!Number.isSafeInteger(n) || n < 0 || n > wq) && xq(), n);
}
function Ik(e, t = "1900") {
  if (typeof e == "number")
    return (
      (!Number.isSafeInteger(e) || e < 0 || e > Xi(t)) &&
        Tk(
          `Date values must be whole Excel ${t} serials between 0 and ${Xi(t)}.`,
        ),
      e
    );
  if (
    ((typeof e != "string" || !/^\d{4}-\d{2}-\d{2}$/.test(e)) &&
      Tk("Enter a date as YYYY-MM-DD or a whole Excel date serial."),
    e === "1900-02-29" && t === "1900")
  )
    return 60;
  let [n, i, r] = e.split("-").map(Number),
    o = kq(n, i - 1, r);
  return (
    (o.getUTCFullYear() !== n ||
      o.getUTCMonth() !== i - 1 ||
      o.getUTCDate() !== r) &&
      Tk("The series date is not a valid calendar date."),
    Ik(_q(o) - ka(t), t)
  );
}
var f8 = {
  delimiter: "auto",
  encoding: "auto",
  quote: '"',
  escape: "double",
  newline: "auto",
  hasHeader: true,
  decimal: ".",
  grouping: "none",
  dateOrder: "ymd",
  empty: "blank",
  trim: false,
  columns: {},
};
var Eq = ["auto", "text", "number", "integer", "boolean", "date"],
  m8 = Array.from({ length: 256 }, (e, t) => String.fromCharCode(t));
[
  8364, 129, 8218, 402, 8222, 8230, 8224, 8225, 710, 8240, 352, 8249, 338, 141,
  381, 143, 144, 8216, 8217, 8220, 8221, 8226, 8211, 8212, 732, 8482, 353, 8250,
  339, 157, 382, 376,
].forEach((e, t) => (m8[128 + t] = String.fromCharCode(e)));
var ao = (e, t, n, i, r) => {
    (e.issueCount++,
      e.issues.length < 100 &&
        e.issues.push({
          record: t,
          column: n,
          message: i,
          ...(r !== undefined ? { value: r.slice(0, 200) } : {}),
        }));
  },
  Sq = (e, t, n, i, r, o, a) => {
    (e.changeCount++,
      e.changes.push({
        record: t,
        column: n,
        address: i,
        reason: r,
        before: o,
        after: a,
      }));
  };
function Cq(e = {}) {
  let t = { ...f8, ...e, columns: { ...(e.columns ?? {}) } };
  for (let [n, i] of Object.entries({
    delimiter: ["auto", ",", ";", "	", "|"],
    encoding: ["auto", "utf-8", "utf-16le", "utf-16be", "windows-1252"],
    quote: ['"', "'", "none"],
    escape: ["double", "backslash"],
    newline: [
      "auto",
      `\r
`,
      `
`,
      "\r",
    ],
    decimal: [".", ","],
    grouping: ["none", ".", ",", " "],
    dateOrder: ["ymd", "dmy", "mdy"],
    empty: ["blank", "text"],
  }))
    if (!i.includes(String(t[n])))
      throw Error(`Unsupported delimited ${n} setting.`);
  if (typeof t.hasHeader != "boolean" || typeof t.trim != "boolean")
    throw Error("Header and whitespace controls must be boolean.");
  if (t.decimal === t.grouping)
    throw Error("Decimal and grouping separators must differ.");
  for (let [n, i] of Object.entries(t.columns))
    if (!/^(0|[1-9]\d*)$/.test(n) || Number(n) >= Ne.maxCols || !Eq.includes(i))
      throw Error("Invalid column type selection.");
  return t;
}
function Tq(e, t, n, i) {
  return {
    schemaVersion: 1,
    direction: e,
    fileName: t,
    fileBytes: n,
    options: i,
    records: 0,
    columns: 0,
    storedCells: 0,
    blankRecords: 0,
    schema: [],
    preview: [],
    issueCount: 0,
    issues: [],
    changeCount: 0,
    changes: [],
    safeTextChanges: 0,
    complete: true,
    valid: false,
    warnings: [],
    safeExportPolicy: "apostrophe-prefix-formula-like-text",
  };
}
function Aq(e, t) {
  if (e.length > Ne.maxFileBytes)
    throw Error("Delimited files must be 10 MB or smaller.");
  let n =
    e[0] === 239 && e[1] === 187 && e[2] === 191
      ? "utf-8"
      : e[0] === 255 && e[1] === 254
        ? "utf-16le"
        : e[0] === 254 && e[1] === 255
          ? "utf-16be"
          : undefined;
  if (n && t !== "auto" && t !== n)
    throw Error(`File BOM indicates ${n}; select that encoding or Automatic.`);
  let i = t === "auto" ? (n ?? "utf-8") : t,
    r;
  try {
    r = (
      i === "windows-1252"
        ? Array.from(e, (o) => m8[o]).join("")
        : new TextDecoder(i, { fatal: true, ignoreBOM: true }).decode(e)
    ).replace(/^\uFEFF/, "");
  } catch {
    throw Error(
      `Invalid ${i} byte sequence. Choose the source encoding; bytes are never silently replaced.`,
    );
  }
  if (r.includes("\0"))
    throw Error(
      "Delimited files cannot contain null bytes; check the encoding.",
    );
  return { text: r, encoding: i };
}
function Iq(e, t, n, i) {
  let r = [],
    o = "",
    a = "start",
    s = 1,
    l = 0,
    c = false,
    u = () => {
      (r.push(o),
        (o = ""),
        (a = "start"),
        r.length > Ne.maxCols &&
          (ao(n, s, undefined, `Delimited data exceeds ${Ne.maxCols} columns.`),
          (c = true)));
    },
    d = (f) => {
      if ((u(), !c)) {
        if (s > Ne.maxRows) {
          (ao(n, s, undefined, `Delimited data exceeds ${Ne.maxRows} rows.`),
            (c = true));
          return;
        }
        ((c = !i(r, s, f === l)), (r = []), s++);
      }
    };
  for (let f = 0; f < e.length && !c; f++) {
    let p = e[f];
    if (a === "quoted")
      t.escape === "backslash" &&
      p === "\\" &&
      (e[f + 1] === t.quote || e[f + 1] === "\\")
        ? (o += e[++f])
        : p === t.quote
          ? t.escape === "double" && e[f + 1] === t.quote
            ? ((o += p), f++)
            : (a = "closed")
          : (o += p);
    else if (p === t.delimiter) u();
    else if (
      p === "\r" ||
      p ===
        `
`
    ) {
      let h =
        p === "\r" &&
        e[f + 1] ===
          `
`
          ? `\r
`
          : p;
      (t.newline !== "auto" &&
        t.newline !== h &&
        ao(
          n,
          s,
          undefined,
          "Record newline does not match the selected dialect.",
        ),
        d(f),
        h.length === 2 && f++,
        (l = f + 1));
    } else
      a === "start" && t.quote !== "none" && p === t.quote
        ? (a = "quoted")
        : (a === "closed"
            ? (ao(n, s, r.length + 1, "Unexpected text after a closing quote."),
              (a = "plain"))
            : t.quote !== "none" &&
              p === t.quote &&
              ao(
                n,
                s,
                r.length + 1,
                "Quote inside an unquoted field; quote the entire field.",
              ),
          (o += p),
          (a = "plain"));
    o.length > dr.maxCellTextLength &&
      (ao(
        n,
        s,
        r.length + 1,
        "A field exceeds the 32,767-character cell limit.",
      ),
      (c = true));
  }
  (!c &&
    a === "quoted" &&
    ao(n, s, r.length + 1, "Quoted field is not closed."),
    !c && l < e.length && d(e.length),
    (n.complete = !c));
}
var Rk = (e) => e.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
function h8(e, t) {
  let n = t.grouping === "none" ? "" : t.grouping,
    i = n ? `(?:\\d+|\\d{1,3}(?:${Rk(n)}\\d{3})+)` : "\\d+";
  if (
    !new RegExp(
      `^[+-]?(?:${i}(?:${Rk(t.decimal)}\\d*)?|${Rk(t.decimal)}\\d+)(?:[eE][+-]?\\d+)?$`,
    ).test(e)
  )
    throw Error("Invalid number for the selected decimal/grouping separators.");
  let o = (n ? e.split(n).join("") : e).replace(t.decimal, "."),
    a = o
      .split(/[eE]/)[0]
      .replace(/[^0-9]/g, "")
      .replace(/^0+/, "")
      .replace(/0+$/, "").length,
    s = Number(o);
  if (
    !Number.isFinite(s) ||
    a > 15 ||
    (s === 0 && /[1-9]/.test(o.split(/[eE]/)[0]))
  )
    throw Error(
      "Number is outside supported precision/range; select Text to retain every digit.",
    );
  return s;
}
function Rq(e, t, n = "1900") {
  let i =
    /^(\d{1,4})([-/])(\d{1,2})\2(\d{1,4})(?:[ T](\d{2}):(\d{2})(?::(\d{2}))?)?$/.exec(
      e,
    );
  if (!i)
    throw Error(
      "Date must match the selected year/month/day order, optionally HH:mm:ss.",
    );
  let r = [i[1], i[3], i[4]],
    [o, a, s] =
      t === "ymd" ? r : t === "dmy" ? [r[2], r[1], r[0]] : [r[2], r[0], r[1]];
  if (o.length !== 4) throw Error("Dates require a four-digit year.");
  let l = Number(i[5] ?? 0),
    c = Number(i[6] ?? 0),
    u = Number(i[7] ?? 0);
  if (l > 23 || c > 59 || u > 59) throw Error("Invalid date/time component.");
  return (
    Ik(`${o}-${a.padStart(2, "0")}-${s.padStart(2, "0")}`, n) +
    (l * 3600 + c * 60 + u) / 86400
  );
}
function Nq(e, t, n, i = "1900") {
  if (e === "") return { raw: n.empty === "text" ? "" : null };
  let r = n.trim ? e.trim() : e;
  if (t === "auto") {
    if (
      /^-?(?:0|[1-9]\d*)(?:\.\d+)?(?:[eE][+-]?\d+)?$/.test(r) &&
      r.split(/[eE]/)[0].replace(/[^0-9]/g, "").length <= 15 &&
      n.decimal === "." &&
      n.grouping === "none"
    )
      try {
        return { raw: h8(r, n) };
      } catch {}
    return { raw: r, ...(r.startsWith("=") ? { literal: true } : {}) };
  }
  if (t === "text")
    return { raw: r, ...(r.startsWith("=") ? { literal: true } : {}) };
  if (t === "date")
    return {
      raw: Rq(r, n.dateOrder, i),
      format: "date",
      numFmt: /[ T]\d\d:/.test(r) ? "yyyy-mm-dd hh:mm:ss" : "yyyy-mm-dd",
    };
  if (t === "boolean") {
    if (/^(true|false)$/i.test(r)) return { raw: r.toLowerCase() === "true" };
    throw Error("Boolean values must be TRUE or FALSE (case-insensitive).");
  }
  let o = h8(r, n);
  if (t === "integer" && !Number.isSafeInteger(o))
    throw Error("Integer column requires a safe whole number.");
  return { raw: o };
}
function g8(e, t, n, i) {
  let r = Cq(n),
    o = Aq(e, r.encoding);
  ((r.encoding = o.encoding),
    r.delimiter === "auto" &&
      (r.delimiter = /\.tsv$/i.test(t)
        ? "	"
        : p8.default.parse(o.text.slice(0, 65536), {
            preview: 10,
            quoteChar: r.quote === "none" ? "" : r.quote,
            escapeChar: r.escape === "double" ? r.quote : "\\",
            delimitersToGuess: [",", ";", "	", "|"],
          }).meta.delimiter || ","));
  let a = Tq("import", t, e.length, r),
    s = {};
  (a.warnings.push(
    "Delimited data contains values only. All imported formula-like content stays literal text and is never treated as a formula. Auto keeps leading-zero/long identifiers and dates as text; select explicit types for localized numbers, dates and booleans.",
  ),
    (n.encoding === undefined || n.encoding === "auto") &&
      a.warnings.push(
        `Encoding: ${r.encoding}; automatic mode uses a BOM when present and UTF-8 otherwise. Legacy encodings require an explicit selection.`,
      ),
    (n.delimiter === undefined || n.delimiter === "auto") &&
      a.warnings.push(
        `Detected delimiter: ${r.delimiter === "	" ? "tab" : r.delimiter}. Review the schema and column count before importing.`,
      ));
  let l = 0;
  if (
    (Iq(o.text, r, a, (c, u, d) => {
      ((a.records = u),
        c.every((f) => f === "") && a.blankRecords++,
        !l && !d && (l = c.length),
        (a.columns = Math.max(a.columns, c.length)),
        !d &&
          l &&
          c.length !== l &&
          ao(
            a,
            u,
            undefined,
            `Malformed row: expected ${l} fields, found ${c.length}.`,
          ));
      for (let f = 0; f < c.length; f++) {
        let p = c[f],
          h = r.hasHeader && u === 1 ? "text" : (r.columns[f] ?? "auto"),
          m = `${Ue(f)}${u}`;
        a.schema[f] ||
          (a.schema[f] = {
            index: f,
            label: r.hasHeader && u === 1 ? p : Ue(f),
            type: r.columns[f] ?? "auto",
            observed: [],
          });
        let g;
        try {
          g = Nq(p, h, r);
        } catch (b) {
          (ao(
            a,
            u,
            f + 1,
            b instanceof Error ? b.message : "Type conversion failed.",
            p,
          ),
            (g = { raw: p, ...(p.startsWith("=") ? { literal: true } : {}) }));
        }
        if (!(r.hasHeader && u === 1)) {
          let b =
            g.raw === null
              ? "blank"
              : g.format === "date"
                ? "date"
                : typeof g.raw;
          a.schema[f].observed.includes(b) || a.schema[f].observed.push(b);
        }
        if (g.raw !== null && !(d && r.empty === "blank")) {
          if (++a.storedCells > Ne.maxCells)
            return (
              ao(
                a,
                u,
                f + 1,
                "Delimited data exceeds 100,000 populated cells.",
              ),
              false
            );
          i && (s[m] = g);
        }
        (p !== "" &&
          (g.raw !== p || (r.trim && p !== p.trim())) &&
          Sq(
            a,
            u,
            f + 1,
            m,
            g.format === "date"
              ? "Date converted to native serial"
              : "Selected type or whitespace conversion",
            p,
            String(g.raw ?? ""),
          ),
          a.preview.length < 25 &&
            f === 0 &&
            a.preview.push({ record: u, raw: c.slice(0, 200), values: [] }),
          u <= 25 && f < 200 && a.preview.at(-1)?.values.push(g.raw));
      }
      return true;
    }),
    a.records ||
      ao(
        a,
        0,
        undefined,
        "The file has no records. Select a populated delimited file.",
      ),
    r.hasHeader && a.records)
  ) {
    let c = a.schema.map((d) => d.label),
      u = new Set();
    c.some(
      (d) =>
        !d.trim() ||
        (u.has(d.toLowerCase()) ? true : (u.add(d.toLowerCase()), false)),
    ) &&
      a.warnings.push(
        "Blank or duplicate header labels are retained exactly. Column letters identify them; no fields are overwritten or renamed.",
      );
  }
  return ((a.valid = a.complete && a.issueCount === 0), { r: a, cells: s });
}
function b8(e, t, n = {}) {
  return g8(e, t, n, false).r;
}
function Eg(e, t, n = {}) {
  let { r, cells } = g8(e, t, n, true);
  if (!r.valid)
    throw Error(
      `Unable to read delimited data: ${r.issues[0]?.message ?? "Incomplete preview."} Review the import schema; no rows were imported.`,
    );
  let o = Q6(
      t
        .replace(/\.[^.]+$/, " ")
        .trim()
        .slice(0, 200) || "Imported workbook",
    ),
    a = o.sheets[0];
  return (
    (a.name = "Imported data"),
    (a.cells = cells),
    (a.frozenRows = r.options.hasHeader ? 1 : 0),
    (a.delimitedSource = {
      records: r.records,
      columns: r.columns,
      hasHeader: r.options.hasHeader,
    }),
    {
      workbook: gi(
        Wm(o, { id: crypto.randomUUID(), name: t.slice(0, 255), kind: "file" }),
      ),
      warnings: r.warnings,
      delimitedReport: r,
    }
  );
}
var Oq = dr.maxXlsxExpandedBytes,
  Mq =
    "XLSX exchange supports cell values, formulas, common formats, supported fonts/strike/alignment/indent/per-edge borders, notes, hidden sheets, column widths, explicit row heights, worksheet default dimensions, bounded hidden rows/columns, frozen rows and columns, basic tables, safe merges, supported validation/conditional-format rules, PNG/JPEG pictures, repeated print titles, full-axis manual print breaks and page setup. XLSX processes at most 100,000 hidden rows per workbook; native JSON keeps larger visibility intervals compactly. Workbook and worksheet-scoped defined names, constant/formula definitions and supported name comments are preserved. Unsupported name attributes are reported; unsafe or malformed definitions stop import/export. Enhanced number formats export supported masks; locale, scaling labels, duration/date-system and accounting alignment fidelity may differ, and enhanced format metadata remains native JSON only. Reusable calculated-column rules are native JSON only; XLSX imports retain supported cell formulas but do not infer rules for future appends. Editable outline groups, threaded discussions and interactive hyperlink metadata remain native-JSON features; XLSX retains only collapsed outline visibility. Native pivot/query refresh definitions, charts/shapes, filters, split panes, default-hidden-row settings, password protection and other unsupported Excel metadata are not preserved. External or executable formulas are converted to literal text on import and export. Keep the original workbook; use native JSON for a lossless application backup.",
  Nk = () => globalThis.crypto.randomUUID();
var $q = (e, t) => ({
    schemaVersion: 1,
    id: Nk(),
    name: e.replace(/\.[^.]+$/, "").slice(0, 200) || "Imported workbook",
    revision: 0,
    metadata: Ip(),
    sheets: t,
    events: [],
  }),
  Pq = (e) =>
    /[\[\]|]/.test(Co(e).replace(/"(?:[^"]|"")*"/g, '""')) ||
    /(?:https?:|file:|javascript:|data:|\\\\)/i.test(e) ||
    /\b(?:_xlfn\.)?(?:WEBSERVICE|HYPERLINK|IMAGE|RTD|DDE|CALL|EXEC|REGISTER(?:\.ID)?|RUN|SQL\.REQUEST|STOCKHISTORY)\s*\(/i.test(
      e,
    );
function Wn(e, t) {
  return new RegExp("(?:^|\\s)" + t + `\\s*=\\s*(["'])([\\s\\S]*?)\\1`).exec(
    e,
  )?.[2];
}
function Dq(e) {
  let t = [];
  for (let n of (e.startsWith("/") ? e.slice(1) : "xl/" + e).split("/"))
    n === ".." ? t.pop() : n !== "." && t.push(n);
  return t.join("/");
}
var Rc = (e) =>
  e.replace(/&(?:amp|lt|gt|quot|apos|#\d+|#x[a-f\d]+);/gi, (t) =>
    t === "&amp;"
      ? "&"
      : t === "&lt;"
        ? "<"
        : t === "&gt;"
          ? ">"
          : t === "&quot;"
            ? '"'
            : t === "&apos;"
              ? "'"
              : String.fromCodePoint(
                  Number.parseInt(
                    t.slice(t[2]?.toLowerCase() === "x" ? 3 : 2, -1),
                    t[2]?.toLowerCase() === "x" ? 16 : 10,
                  ),
                ),
  );
function v8(e, t) {
  let n = vg(e),
    i = new Set(n.keys()),
    r = n.size,
    o = ON();
  if (
    [...i].some((_) =>
      /(?:vbaProject\.bin|xl\/macrosheets\/|xl\/activeX\/|xl\/embeddings\/)/i.test(
        _,
      ),
    )
  )
    throw new Error(
      "Workbooks containing macros, ActiveX or embedded objects are not accepted. Save a clean XLSX copy first.",
    );
  let a = 0,
    s,
    l = 0,
    c = 0,
    u = 0,
    d = 0,
    f = 0,
    p = 0,
    h = 0,
    m = false,
    g = new Set(),
    b = new gc((_) => {
      if ((l++, !i.has(_.name) || g.has(_.name))) {
        s = new Error("Archive entries do not match its directory.");
        return;
      }
      g.add(_.name);
      let S = 0,
        O = /\.(?:xml|rels|vml)$/i.test(_.name),
        E = O ? new TextDecoder("utf-8", { fatal: true }) : undefined,
        R = -1,
        I = "";
      ((_.ondata = (U, P, T) => {
        if (U) {
          s = new Error(
            `Corrupt XLSX compressed content in ${_.name}. Repair the original workbook and save a clean XLSX copy.`,
          );
          return;
        }
        if (((a += P.length), (S += P.length), a > Oq)) {
          (_.terminate(),
            (s = new Error("The XLSX archive expands beyond 50 MB.")));
          return;
        }
        if (S > n.get(_.name).size) {
          (_.terminate(),
            (s = new Error(
              "The XLSX archive expansion does not match its directory.",
            )));
          return;
        }
        R = Tc(R, P);
        try {
          E && (I += E.decode(P, { stream: !T }));
        } catch {
          s = new Error(
            `Unsupported XML encoding in ${_.name}. Save a standard UTF-8 XLSX copy.`,
          );
          return;
        }
        if (T) {
          if (S !== n.get(_.name).size) {
            s = new Error(
              "The XLSX archive expansion does not match its directory.",
            );
            return;
          }
          if ((R ^ -1) >>> 0 !== n.get(_.name).crc) {
            s = new Error(
              `XLSX checksum mismatch in ${_.name}. The archive is corrupt; repair the original and save a clean XLSX copy.`,
            );
            return;
          }
          try {
            if ((O && MN(I, _.name, n, o), /^xl\/styles\.xml$/i.test(_.name)))
              for (let A of I.matchAll(/<(?:\w+:)?sz\b([^>]*)>/g)) {
                let L = Number(Wn(A[1], "val"));
                (!Number.isInteger(L) || L < 8 || L > 72) && (m = true);
              }
            if (/^xl\/worksheets\//i.test(_.name)) {
              (l8(I, _.name),
                (h = RN(I, _.name, h)),
                /^xl\/worksheets\/[^/]+\.xml$/i.test(_.name) &&
                  t.numericTokens.set(_.name, XN(I)));
              for (let v of I.matchAll(/<(?:\w+:)?pageSetup\b([^>]*)>/gi)) {
                let k = Wn(v[1], "pageOrder");
                if (
                  k !== undefined &&
                  !["overThenDown", "downThenOver"].includes(Rc(k))
                )
                  throw new Error(
                    "Unsupported XLSX page order. Import stopped to preserve the source layout.",
                  );
              }
              let A = {};
              for (let [v, k, x, C] of [
                ["rowBreaks", "rowBreaks", Ne.maxRows, Ne.maxCols],
                ["colBreaks", "columnBreaks", Ne.maxCols, Ne.maxRows],
              ]) {
                let N = [
                  ...I.matchAll(
                    new RegExp(
                      "<(?:\\w+:)?" +
                        v +
                        "\\b[^>]*>([\\s\\S]*?)</(?:\\w+:)?" +
                        v +
                        ">",
                      "gi",
                    ),
                  ),
                ];
                if (N.length > 1)
                  throw Error(
                    "Duplicate XLSX page-break sections are not supported.",
                  );
                let F = [],
                  z = 0;
                for (let q of (N[0]?.[1] ?? "").matchAll(
                  /<(?:\w+:)?brk\b([^>]*)\/?\s*>/gi,
                )) {
                  if (++z > 100)
                    throw Error(
                      "XLSX page breaks exceed 100 boundaries per axis.",
                    );
                  let D = Wn(q[1], "id"),
                    W = Wn(q[1], "man");
                  if (
                    !/^[1-9]\d*$/.test(D ?? "") ||
                    !Number.isSafeInteger(Number(D)) ||
                    Number(D) >= x
                  )
                    throw Error("Invalid XLSX page-break boundary.");
                  if (W !== "1" && W !== "true") continue;
                  let $ = Number(Wn(q[1], "min") ?? 0),
                    M = Number(Wn(q[1], "max") ?? C - 1);
                  if ($ !== 0 || !Number.isSafeInteger(M) || M < C - 1)
                    throw Error(
                      "Partial-span XLSX page breaks are not supported. Use full-row or full-column breaks.",
                    );
                  F.push(Number(D));
                }
                F.sort((q, D) => q - D);
                let j = ug(F, x);
                j && (A[k] = j);
              }
              (t.printBreaks.set(_.name, A),
                t.conditionalStops.set(_.name, JN(I)),
                t.validationAreas.set(_.name, ZN(I)));
              let L = [],
                y = [];
              for (let v of I.matchAll(/<(?:\w+:)?col\b([^>]*)>/gi)) {
                let k = Wn(v[1], "min"),
                  x = Wn(v[1], "max"),
                  C = Wn(v[1], "width"),
                  N = Number(k) - 1,
                  F = Number(x) - 1;
                if (
                  !/^[1-9]\d*$/.test(k ?? "") ||
                  !/^[1-9]\d*$/.test(x ?? "") ||
                  !Number.isSafeInteger(N) ||
                  !Number.isSafeInteger(F) ||
                  N < 0 ||
                  F < N ||
                  F >= Ne.maxCols ||
                  y.length >= Ne.maxCols
                )
                  throw Error(
                    "XLSX column metadata must stay within 16,384 nonoverlapping columns.",
                  );
                let z = C === undefined ? undefined : Number(C);
                if (
                  z !== undefined &&
                  (!C?.trim() || !Number.isFinite(z) || z < 0)
                )
                  throw Error(
                    "XLSX column widths must be finite nonnegative values.",
                  );
                y.push({
                  start: N,
                  end: F,
                  ...(z === undefined ? {} : { width: z }),
                });
              }
              y.sort((v, k) => v.start - k.start);
              for (let v = 1; v < y.length; v++)
                if (y[v].start <= y[v - 1].end)
                  throw Error(
                    "XLSX column metadata must stay within 16,384 nonoverlapping columns.",
                  );
              if (
                (t.columns.set(_.name, y),
                /<(?:\w+:)?sheetFormatPr\b[^>]*\bzeroHeight\s*=\s*["'](?:1|true)["']/i.test(
                  I,
                ))
              )
                throw new Error(
                  "Default-hidden worksheet rows cannot be imported safely. Replace the default with explicit hidden rows before importing.",
                );
              for (let v of I.matchAll(/<(?:\w+:)?row\b([^>]*)>/gi)) {
                if (!/\bhidden\s*=\s*["'](?:1|true)["']/i.test(v[1])) continue;
                let k = Number(Wn(v[1], "r")) - 1;
                if (!Number.isInteger(k) || k < 0 || k >= Ne.maxRows)
                  throw new Error(
                    "Hidden XLSX rows must have valid row numbers.",
                  );
                let x = L.at(-1);
                if (x && k <= x.end)
                  throw new Error(
                    "Hidden XLSX row numbers must be distinct and ascending.",
                  );
                if (
                  (x && k === x.end + 1
                    ? (x.end = k)
                    : L.push({ start: k, end: k }),
                  L.length > Na.maxIntervalsPerAxis)
                )
                  throw new Error(
                    "Hidden metadata exceeds 10,000 intervals per worksheet axis.",
                  );
                if (++p > Na.maxXlsxHiddenRows)
                  throw new Error(
                    "XLSX hidden rows exceed the 100,000-row processing budget. Reduce hidden rows in the source; native JSON preserves larger visibility intervals.",
                  );
              }
              L.length && t.hiddenRows.set(_.name, L);
            }
            if (_.name === "xl/workbook.xml")
              for (let A of I.matchAll(/<(?:\w+:)?sheet\b([^>]*)>/g)) {
                let L = Wn(A[1], "name"),
                  y = Wn(A[1], "r:id");
                L && y && t.sheetNames.set(Rc(L), Rc(y));
              }
            if (_.name === "xl/_rels/workbook.xml.rels")
              for (let A of I.matchAll(/<(?:\w+:)?Relationship\b([^>]*)>/g)) {
                let L = Wn(A[1], "Id"),
                  y = Wn(A[1], "Target");
                L &&
                  y &&
                  Wn(A[1], "TargetMode") !== "External" &&
                  t.sheetRelationships.set(Rc(L), Dq(Rc(y)));
              }
            for (let A of I.matchAll(
              /<(?:\w+:)?(mergeCell|dataValidation)\s[^>]*?\b(ref|sqref)\s*=\s*["']([^"']+)["'][^>]*>/g,
            ))
              for (let L of Rc(A[3]).trim().split(/\s+/)) {
                let y = Qe(L),
                  v =
                    (y.lastRow - y.firstRow + 1) * (y.lastCol - y.firstCol + 1);
                if (A[1] === "mergeCell") {
                  if (v > 500 || ++f > 1e3 || (u += v) > Ne.maxCells)
                    throw new Error(
                      "Merged ranges exceed the supported import budget (500 cells per merge and 100,000 cells overall).",
                    );
                } else if ((d += v) > Ne.maxCells)
                  throw new Error(
                    "XLSX validation ranges exceed the 100,000-cell import budget. Narrow the validation area in the source workbook.",
                  );
              }
            if (_.name === "xl/workbook.xml") {
              t.calculation = e8(I);
              let A = o5(I);
              ((t.names = A.names),
                (t.nameWarnings = A.warnings),
                (t.builtinNameXml = A.builtinXml));
            }
          } catch (A) {
            s =
              A instanceof Error
                ? A
                : new Error("Invalid XLSX worksheet metadata.");
            return;
          }
          c++;
        }
      }),
        _.start());
    });
  b.register(mc);
  for (let _ = 0; _ < e.length; _ += 2048) {
    if (s) throw s;
    try {
      b.push(e.subarray(_, Math.min(_ + 2048, e.length)), _ + 2048 >= e.length);
    } catch {
      throw new Error(
        "Corrupt XLSX compressed content. Repair the original workbook and save a clean XLSX copy.",
      );
    }
  }
  if (s) throw s;
  if (l !== r || c !== r)
    throw new Error("The XLSX archive contains incomplete entries.");
  let w = [Mq, ...t.nameWarnings, ...$N(o, i)];
  return (
    (t.emptyStringCaches = o.emptyStringCaches),
    m &&
      w.push(
        "Unsupported fractional or out-of-range font sizes were found in the source. The XLSX reader may normalize them; the editor supports integer font sizes from 8 to 72 points.",
      ),
    [...i].some((_) => /externalLinks|connections\.xml/i.test(_)) &&
      w.push(
        "External workbook links or connections were detected. They will not be refreshed or included in the derived copy; the original retains their source parts.",
      ),
    w
  );
}
function Lq(e) {
  if (!(!e || e === "General"))
    return /%/.test(e)
      ? "percent"
      : /[$€£¥]/.test(e)
        ? "currency"
        : /[ymdhis]/i.test(e.replace(/"[^"]*"/g, ""))
          ? "date"
          : "number";
}
function Ga(e) {
  return e && /^(?:[0-9a-f]{8}|[0-9a-f]{6})$/i.test(e)
    ? `#${e.slice(-6).toLowerCase()}`
    : undefined;
}
function Uq(e, t = [], n, i = false) {
  let r = (m) => {
      t.includes(m) || t.push(m);
    },
    o = e.value;
  if (o === null && !e.note && !Object.keys(e.style ?? {}).length) return;
  let a = null,
    s;
  typeof o == "string"
    ? ((a = o), a.startsWith("=") && (s = true))
    : typeof o == "number" || typeof o == "boolean"
      ? (a = o)
      : o instanceof Date
        ? (a = n !== undefined ? Number(n) : _p(o, i ? "1904" : "1900"))
        : o &&
          typeof o == "object" &&
          ("formula" in o || "sharedFormula" in o
            ? ((a = `=${e.formula}`), Pq(a) && (s = true))
            : "richText" in o
              ? (r(
                  "Rich-text runs are flattened to cell text; per-run formatting is not preserved.",
                ),
                (a = o.richText.map((m) => m.text).join("")),
                a.startsWith("=") && (s = true))
              : "hyperlink" in o
                ? ((a = o.text), a.startsWith("=") && (s = true))
                : "error" in o && (a = o.error));
  let l = {};
  (e.font?.bold && (l.bold = true),
    e.font?.italic && (l.italic = true),
    e.font?.strike && (l.strikethrough = true),
    e.alignment?.vertical &&
      ["top", "middle", "bottom"].includes(e.alignment.vertical) &&
      (l.verticalAlign = e.alignment.vertical),
    e.alignment?.indent !== undefined &&
      (Number.isInteger(e.alignment.indent) &&
      e.alignment.indent >= 0 &&
      e.alignment.indent <= 15
        ? (l.indent = e.alignment.indent)
        : r(
            "Unsupported indentation was omitted; the editor supports integer levels 0\u201315.",
          )),
    e.font?.underline && (l.underline = true),
    Ca.includes(e.font?.name) && (l.fontFamily = e.font.name),
    e.font?.size !== undefined &&
      Number.isInteger(e.font.size) &&
      e.font.size >= 8 &&
      e.font.size <= 72 &&
      (l.fontSize = e.font.size),
    e.alignment?.wrapText && (l.wrap = true));
  let c = {};
  for (let m of Ts) {
    let g = e.border?.[m];
    if (!g?.style) continue;
    if (!Nw.includes(g.style)) {
      r(
        "An unsupported border line was omitted; supported edges and base values were retained.",
      );
      continue;
    }
    let b = Ga(g.color?.argb);
    (g.color &&
      !b &&
      r(
        "A theme/indexed border color used a black fallback; theme bindings are not preserved.",
      ),
      (c[m] = { style: g.style, color: b ?? "#000000" }));
  }
  let u = Object.keys(c),
    d = u.every(
      (m) => c[m]?.style === "thin" && c[m]?.color?.toLowerCase() === "#000000",
    );
  (d && u.length === 4
    ? (l.border = "all")
    : d && u.length === 1 && c.bottom
      ? (l.border = "bottom")
      : u.length && (l.borders = c),
    e.border?.diagonal &&
      r("Diagonal borders are unsupported and were omitted."),
    e.font?.name &&
      !Ca.includes(e.font.name) &&
      r(
        "An unsupported font family uses the editor default; exact font fidelity is not preserved.",
      ),
    e.font?.size !== undefined &&
      (!Number.isInteger(e.font.size) || e.font.size < 8 || e.font.size > 72) &&
      r(
        "An unsupported font size was omitted; supported sizes are integers from 8 to 72 points.",
      ),
    e.font?.underline &&
      e.font.underline !== true &&
      e.font.underline !== "single" &&
      r("An unsupported underline variant was simplified to single underline."),
    (e.font?.vertAlign ||
      e.alignment?.textRotation ||
      e.alignment?.shrinkToFit ||
      (e.alignment?.horizontal &&
        !["left", "center", "right"].includes(e.alignment.horizontal)) ||
      (e.alignment?.vertical &&
        !["top", "middle", "bottom"].includes(e.alignment.vertical))) &&
      r(
        "Unsupported text positioning/alignment settings were omitted; raw values were retained.",
      ),
    e.font?.color &&
      !Ga(e.font.color.argb) &&
      r(
        "A theme/indexed font color uses the editor default; theme bindings are not preserved.",
      ),
    e.fill &&
      (e.fill.type !== "pattern" ||
        !["solid", "none"].includes(e.fill.pattern)) &&
      r(
        "An unsupported gradient or pattern fill was omitted; base cell values were retained.",
      ),
    e.fill?.type === "pattern" &&
      e.fill.pattern === "solid" &&
      e.fill.fgColor &&
      !Ga(e.fill.fgColor.argb) &&
      r(
        "A theme/indexed fill color uses the editor default; theme bindings are not preserved.",
      ));
  let f = Ga(e.font?.color?.argb);
  if (
    (f && (l.color = f),
    e.fill?.type === "pattern" && e.fill.pattern === "solid")
  ) {
    let m = Ga(e.fill.fgColor?.argb);
    m && (l.fill = m);
  }
  e.alignment &&
    ["left", "center", "right"].includes(e.alignment.horizontal ?? "") &&
    (l.align = e.alignment.horizontal);
  let p =
      typeof e.note == "string"
        ? e.note
        : e.note?.texts?.map((m) => m.text).join(""),
    h = e.numFmt && e.numFmt !== "General" ? e.numFmt : undefined;
  return {
    raw: a,
    ...(s ? { literal: s } : {}),
    ...(Object.keys(l).length ? { style: l } : {}),
    ...(h ? { numFmt: h, format: Lq(h) } : {}),
    ...(p ? { note: p } : {}),
  };
}
async function kg(e) {
  if (/\.xltx$/i.test(e.name)) return zN(e);
  if (/\.json$/i.test(e.name)) {
    if (e.size > Ne.maxNativeBytes)
      throw new Error("Native workbook exceeds the 50 MB import limit.");
    let S = Fq(await e.text());
    return { workbook: S, warnings: [...B6(S), ...Bw(S)] };
  }
  if (e.size > Ne.maxFileBytes)
    throw new Error("Files must be 10 MB or smaller.");
  if (/\.(xlsm|xlsb|xls|xltm|xlam)$/i.test(e.name))
    throw new Error(
      "Legacy or macro-enabled Excel formats are not supported. Save as XLSX, CSV or native JSON.",
    );
  if (/\.(csv|tsv)$/i.test(e.name))
    return Eg(new Uint8Array(await e.arrayBuffer()), e.name, {
      hasHeader: false,
    });
  if (!/\.xlsx$/i.test(e.name))
    throw new Error("Choose an XLSX, CSV, TSV or JSON file.");
  let t = new Uint8Array(await e.arrayBuffer()),
    n = {
      names: [],
      nameWarnings: [],
      numericTokens: new Map(),
      builtinNameXml: "",
      hiddenRows: new Map(),
      conditionalStops: new Map(),
      validationAreas: new Map(),
      printBreaks: new Map(),
      columns: new Map(),
      sheetRelationships: new Map(),
      sheetNames: new Map(),
    };
  if (t.byteLength > Ne.maxFileBytes)
    throw new Error("XLSX files must be 10 MB or smaller.");
  let i = v8(t, n),
    r = ag(t, e.name),
    o = xk(Nd(t), r.sha256),
    a = new y8.default.Workbook(),
    s = t;
  if (n.names.length || n.nameWarnings.length) {
    let S = Nd(t);
    ((S["xl/workbook.xml"] = za(
      a5(Ao(S["xl/workbook.xml"]), n.builtinNameXml),
    )),
      (s = Rd(S)));
  }
  try {
    await a.xlsx.load(s);
  } catch {
    throw new Error(
      "The XLSX workbook contains corrupt or unsupported records. Open the original in Excel, repair it if needed, and save a clean XLSX copy.",
    );
  }
  if (
    a.worksheets.length !== n.sheetNames.size ||
    a.worksheets.some((S) => !n.sheetNames.has(S.name))
  )
    throw new Error(
      "Unsupported or missing worksheet records. Save a standard XLSX copy with ordinary worksheets.",
    );
  if (!a.worksheets.length)
    throw new Error("A workbook must have at least one sheet.");
  let l = 0,
    c = 0,
    u = a.worksheets.map((S) => {
      let O = Math.max(S.columnCount, S.columns?.length ?? 0);
      if (S.rowCount > Ne.maxRows || O > Ne.maxCols)
        throw new Error(
          `Worksheet ${S.name} exceeds ${Ne.maxRows} rows or ${Ne.maxCols} columns.`,
        );
      let E = {},
        R = S.model,
        I = {},
        U = n.sheetRelationships.get(n.sheetNames.get(S.name) ?? ""),
        P = n.hiddenRows.get(U ?? "") ?? [],
        T = (N, F) => {
          let z = N.at(-1);
          if (z && F === z.end + 1) z.end = F;
          else {
            if (N.length >= Na.maxIntervalsPerAxis)
              throw new Error(
                "Hidden metadata exceeds 10,000 intervals per worksheet axis.",
              );
            N.push({ start: F, end: F });
          }
        };
      for (let N of R.rows) {
        if (typeof N.height == "number" && Number.isFinite(N.height)) {
          let F = (N.height * 4) / 3;
          ((I[String(N.number - 1)] = Math.max(21, Math.min(546, F))),
            (F < 21 || F > 546) && c++);
        }
        for (let F of N.cells) {
          let z = S.getCell(F.address);
          if (z.isMerged && z.master.address !== z.address) continue;
          let j = Uq(
            z,
            i,
            n.numericTokens.get(U ?? "")?.get(F.address),
            !!a.properties.date1904,
          );
          if (j) {
            if (++l > Ne.maxCells)
              throw new Error("Workbook exceeds 100,000 populated cells.");
            E[z.address] = j;
          }
        }
      }
      let A = S.views?.find((N) => N.state === "frozen"),
        L = [];
      (S.columns ?? []).forEach((N, F) => {
        N.hidden && T(L, F);
      });
      let y = {};
      for (let N of n.columns.get(U ?? "") ?? []) {
        if (N.width === undefined) continue;
        let F = N.width * 7;
        (F < 40 || F > 600) &&
          i.push(
            "Worksheet " +
              S.name +
              " column widths were clamped to the supported pixel bounds.",
          );
        let z = Math.max(40, Math.min(600, F));
        for (let j = N.start; j <= N.end; j++) y[j] = z;
      }
      let v = (N, F, z, j, q) => {
          if (N === undefined || N === 0) return;
          if (typeof N != "number" || !Number.isFinite(N) || N <= 0)
            throw Error("Invalid worksheet default " + q + ".");
          let D = N * F;
          return (
            (D < z || D > j) &&
              i.push(
                "Worksheet " +
                  S.name +
                  " default " +
                  q +
                  " was clamped to the supported pixel bounds.",
              ),
            Math.max(z, Math.min(j, D))
          );
        },
        k = v(S.properties.defaultRowHeight, 4 / 3, 21, 546, "row height"),
        x = v(S.properties.defaultColWidth, 7, 40, 600, "column width"),
        C = {
          id: Nk(),
          name: S.name,
          ...(k !== undefined ? { defaultRowHeight: k } : {}),
          ...(x !== undefined ? { defaultColumnWidth: x } : {}),
          cells: E,
          rowCount: Ne.maxRows,
          colCount: Ne.maxCols,
          frozenRows: A?.ySplit ?? 0,
          ...(A?.xSplit ? { frozenCols: A.xSplit } : {}),
          ...(P.length ? { hiddenRows: P } : {}),
          ...(L.length ? { hiddenCols: L } : {}),
          ...(Object.keys(y).length ? { columnWidths: y } : {}),
          ...(Object.keys(I).length ? { rowHeights: I } : {}),
          ...(S.state !== "visible" ? { hidden: true } : {}),
          ...(S.state === "veryHidden" ? { veryHidden: true } : {}),
          ...(Ga(S.properties.tabColor?.argb)
            ? { tabColor: Ga(S.properties.tabColor?.argb) }
            : {}),
        };
      (i.push(...j5(S, C)),
        i.push(
          ...YN(
            S,
            C,
            n.conditionalStops.get(U ?? ""),
            n.validationAreas.get(U ?? ""),
          ),
        ),
        i.push(...SN(a, S, C, { containerFilename: e.name })));
      try {
        lN(S, C);
        let N = n.printBreaks.get(U ?? "");
        N && Wd({ ...C.pageLayout, ...N }, C);
      } catch (N) {
        i.push(`Page setup on ${C.name} could not be retained: ${N.message}`);
      }
      return C;
    });
  (c &&
    i.push(
      `${c} explicit row height(s) were adjusted to the supported ${21}\u2013${546} CSS-pixel range.`,
    ),
    a.properties.date1904 &&
      i.push(
        "The source uses the 1904 date system. Its date system, numeric serials and formulas are preserved; calendar dates are not silently converted to 1900.",
      ),
    i.push(
      "Formula results are recalculated by the supported local formula engine. Unsupported Excel functions may display errors. Cell error values are imported as displayed text. Audit history and application tickmarks are available only in native JSON backups.",
    ));
  let d = $q(e.name, u);
  d.xlsxSource = r;
  let f = FN(Nd(t), r.sha256);
  (f.length &&
    (d.workbookLinks = { schemaVersion: 1, sources: [], packageTargets: f }),
    a.properties.date1904 && (d.datePolicy = { ...Hl, dateSystem: "1904" }));
  let p = 0,
    h = [];
  for (let [S, O] of n.sheetNames)
    for (let [E, R] of n.numericTokens.get(n.sheetRelationships.get(O) ?? "") ??
      []) {
      let I = VN(R);
      I && (p++, h.length < 8 && h.push(`${S}!${E}: ${R.slice(0, 80)} (${I})`));
    }
  (p &&
    i.push(
      `Numeric precision risk: ${p} numeric value/cache(s) may lose significant digits in Excel or binary conversion. Numeric cells retain their declared numeric type; store identifiers as Text and keep the source file. ${h.join("; ")}`,
    ),
    a.created instanceof Date &&
      Number.isFinite(a.created.getTime()) &&
      a.modified instanceof Date &&
      Number.isFinite(a.modified.getTime()) &&
      a.modified >= a.created &&
      (d.metadata = {
        createdAt: a.created.toISOString(),
        modifiedAt: a.modified.toISOString(),
      }));
  let m = s5(n.names, d, [...n.sheetNames.keys()]);
  m.length && (d.definedNames = m);
  let g = n.calculation ?? {
    mode: "auto",
    iterative: false,
    maxIterations: null,
    maxChange: null,
  };
  t8(d, g, i);
  let b = gi(Wm(d, { id: Nk(), name: e.name, kind: "file" })),
    w = _k("import", e.name, t.byteLength, b, i);
  ((w.partInventory = o),
    w.warnings.push(
      "The exact source XLSX is retained in this workbook and native JSON. Passive custom metadata can be copied unchanged; all other unsupported parts remain accessible in the original. Save native JSON to retain that source across sessions.",
    ));
  let _ = {
    schemaVersion: 1,
    fileName: e.name,
    sourceSha256: Array.from(
      new Uint8Array(
        await crypto.subtle.digest("SHA-256", new Uint8Array(t).buffer),
      ),
    )
      .map((S) => S.toString(16).padStart(2, "0"))
      .join(""),
    importedAt: new Date().toISOString(),
    sourceSettings: g,
    entries: [],
  };
  return (
    a8(
      a,
      b,
      w,
      new Map(
        [...n.sheetNames].map(([S, O]) => [
          S,
          n.emptyStringCaches?.get(n.sheetRelationships.get(O) ?? "") ??
            new Set(),
        ]),
      ),
      _,
      new Map(
        [...n.sheetNames].map(([S, O]) => [
          S,
          n.numericTokens.get(n.sheetRelationships.get(O) ?? "") ?? new Map(),
        ]),
      ),
    ),
    (b.importedCalculationCaches = _),
    w.warnings.push(...Bw(b)),
    { workbook: gi(b), warnings: w.warnings, report: w }
  );
}
function Fq(e) {
  if (new TextEncoder().encode(e).length > Ne.maxNativeBytes)
    throw new Error("Native workbook exceeds the 50 MB import limit.");
  let t;
  try {
    t = JSON.parse(e);
  } catch {
    throw new Error("The native workbook is not valid JSON.");
  }
  return gi(t);
}
function zq() {
  return {
    names: [],
    nameWarnings: [],
    numericTokens: new Map(),
    builtinNameXml: "",
    hiddenRows: new Map(),
    conditionalStops: new Map(),
    validationAreas: new Map(),
    printBreaks: new Map(),
    columns: new Map(),
    sheetRelationships: new Map(),
    sheetNames: new Map(),
  };
}
function w8(e) {
  let t = XR(e),
    n = v8(t.bytes, zq()),
    i = _k("import", t.fileName, t.bytes.length, e, n);
  return ((i.partInventory = xk(Nd(t.bytes), t.sha256)), i);
}
var x8 = [
    "#BLOCKED!",
    "#CALC!",
    "#CYCLE!",
    "#DIV/0!",
    "#ERROR!",
    "#LIMIT!",
    "#LIMIT! Calculation inputs exceed the working-data budget.",
    "#LIMIT! Calculation result payload exceeds the working-data budget.",
    "#N/A",
    "#NAME?",
    "#NUM!",
    "#PENDING!",
    "#REF!",
    "#SPILL!",
    "#TIMEOUT!",
    "#VALUE!",
    "1,048,576",
    "100,000",
    "100000",
    "1048576",
    "16,384",
    "16384",
    "32,767",
    "32767",
    "A DAX measure requires an expression and no aggregate column.",
    "A Goal Seek cell is outside Excel's address bounds.",
    "A Goal Seek source is outside the visible, authorized host context.",
    "A LAMBDA must be called with arguments or saved as a reusable name.",
    "A PDF page did not finish within 30 seconds. Reduce its complexity or use browser printing.",
    "A Solver worksheet reference is missing.",
    "A blank source cell could be a spill child of a formula outside the source. Use a reviewed values snapshot or include and validate the scalar source formula.",
    "A bounded reviewer name is required.",
    "A calculated column fill supports at most 500 data-body rows.",
    "A calculated field needs add/subtract/multiply/divide and exactly one numeric constant or other source field.",
    "A cash-flow date precedes the first cash-flow date.",
    "A cash-flow or security date is missing or is not a date.",
    "A cell must have a raw value.",
    "A changing worksheet is protected.",
    "A changing worksheet is protected. Unprotect it before analyzing scenarios.",
    "A character cannot be represented in Windows-1252. Choose UTF-8; text is never replaced.",
    "A chart filter row is outside its source range. Reset the row filter.",
    "A chart source cannot follow data to another worksheet in this implementation. Move the chart data on its current sheet.",
    "A checkpoint belongs to another workbook; original storage was preserved.",
    "A column condition must use its own column.",
    "A column conversion refers to a field outside the split result.",
    "A compatibility report is required for format conversion.",
    "A conditional formatting rule requires an applied range.",
    "A copied analysis refers to a missing source table.",
    "A current-row structured reference cannot be safely moved out of its table context.",
    "A custom rule requires a formula relative to the range's top-left cell.",
    "A data row or column cannot fit beside the repeated titles. Reduce the scale or title dimensions.",
    "A dependent table/header rewrite would leave the approved scope. Migrate a separate range instead.",
    "A derived constant must be bounded text or a finite number.",
    "A destination worksheet is outside the host write scope.",
    "A drawing crosses a repeated-title boundary. Move or resize the drawing, or adjust the print titles first.",
    "A dynamic formula could refer to the moving totals row. Preserve or remove that ambiguity before appending.",
    "A failed calculation cannot be applied. The baseline is unchanged.",
    "A failed recipe cannot publish output. Repair the failed step and preview again.",
    "A field exceeds the 32,767-character cell limit.",
    "A field needs quoting; choose a quote character to preserve delimiters/newlines.",
    "A fill or paste operation may change at most 500 cells.",
    "A filter cannot reach any selected measure. Connect its table or remove the filter.",
    "A formula dependency is outside the visible, authorized host context.",
    "A formula refers directly to the moving totals row. Use a structured table totals reference or move those references before appending.",
    "A formula rule requires a formula beginning with =.",
    "A formula watch address is outside Excel's cell bounds.",
    "A formula watch must identify one valid Excel cell address.",
    "A generated pivot/query output cannot be the Goal Seek variable.",
    "A join header cannot be repeated in the key pairs.",
    "A list rule needs nonempty values.",
    "A local evidence picture requires its original filename.",
    "A local workbook namespace must contain 1\u2013300 characters.",
    "A marked calendar requires a nonnullable physical date column.",
    "A merged area crosses a repeated-title boundary. Unmerge it or adjust the print titles first.",
    "A merged area supports at most 500 cells.",
    "A merged cell cannot be the Goal Seek variable.",
    "A migrated formula has invalid syntax, a circular reference, an unsupported reference or a calculation-limit failure.",
    "A missing named style must use an empty fallback base.",
    "A missing theme must use the Office fallback.",
    "A model table requires at least one physical source column.",
    "A model table supports at most 500 data rows, 50 fields and 2,000 cells including headers.",
    "A move may change at most 500 cells in total, including source clears, destination cells, and dependent references. Move a smaller block.",
    "A name already exists in this scope.",
    "A name must refer to a formula beginning with = (up to 4,096 characters).",
    "A name operation may affect at most 500 dependent locations. Split the operation into smaller workbooks.",
    "A name operation may rewrite at most 500 formulas or definitions. Split the migration into smaller workbooks.",
    "A named style base cannot contain another binding.",
    "A named style definition needs supported formatting without another binding.",
    "A nil XML field cannot contain a value.",
    "A one-variable table needs one input and a row or column orientation.",
    "A private calculation context could not be isolated.",
    "A quote must begin a field.",
    "A quote must begin a field; escape embedded quotes by doubling them inside a quoted field.",
    "A quoted field is not closed. Check the source cell before splitting.",
    "A recipe source is outside the authorized host context.",
    "A recipe step exceeds 1,000 rows or 20,000 cells. Narrow the source; no rows were truncated.",
    "A recipe with this name already exists.",
    "A reference exceeds Excel's row boundary.",
    "A reference intersects hidden cells; their content was not inspected.",
    "A reference points outside the permitted visible worksheet context; target details are withheld.",
    "A referenced range intersects hidden rows or columns; hidden cell details are excluded.",
    "A referenced worksheet is missing or outside the current context.",
    "A referenced worksheet is outside the selected inspection context; its cells and values were excluded.",
    "A report snapshot exceeds 4 MB.",
    "A report source cell is missing, hidden or outside the authorized context.",
    "A report with this name already exists.",
    "A restore actor must contain 1\u2013100 characters.",
    "A result exceeds the 32,767-character cell limit.",
    "A review decision refers to an unknown cell.",
    "A review thread with a missing worksheet must be retained as orphaned.",
    "A reviewed column type operation supports at most 500 data rows. Preserve the native source and use bounded extracts for larger conversions.",
    "A reviewed table operation may change at most 500 cells.",
    "A row permutation must contain every destination exactly once within its source block.",
    "A saved analysis already has this name. Choose another name or update that analysis.",
    "A saved checkpoint is corrupt; original storage was preserved.",
    "A saved source worksheet was deleted. This definition cannot reconnect by name. Remove the definition or create a new analysis with current source cells; its historical report is retained.",
    "A saved view must keep its source worksheet.",
    "A saved workbook needs a visible worksheet.",
    "A scalar argument must contain exactly one value.",
    "A screened XLSX conversion and compatibility report are required.",
    "A selected Goal Seek worksheet no longer exists.",
    "A selected chart series no longer exists in its source range. Re-select series.",
    "A selected formula or dependency has a calculation error. Correct it before screening.",
    "A selected model measure is missing.",
    "A selected range contains hidden content. Explicitly include permitted hidden content or choose a visible range.",
    "A selected source is outside the authorized comparison context.",
    "A selected watch no longer exists. Review the current watch list.",
    "A selected watch source is outside the visible, authorized context. Review visibility or host access before adding it.",
    "A sensitivity table supports at most 400 combinations.",
    "A series may contain at most 500 target cells.",
    "A source calculation attempted an unavailable or unauthorized dependency.",
    "A source cell contains more than 50 fields. Choose a different delimiter or split a smaller record.",
    "A source cell must contain text of at most 32,767 characters.",
    "A source contains a calculation error, nonfinite number or oversized text. Repair it before capturing source values.",
    "A source encountered an operational calculation failure, including an error hidden by a formula handler.",
    "A source formula has incomplete, cyclic, hidden, unauthorized or unsupported dependencies.",
    "A source header conflicts with a generated provenance column name.",
    "A source includes hidden cells. Explicitly include hidden records before previewing.",
    "A source is hidden. Explicitly include hidden cells within the authorized scope.",
    "A source is outside the authorized host context.",
    "A source note/formula exceeds the drill-output note budget.",
    "A source rectangle is reversed or outside Excel's address bounds.",
    "A source region exceeds the declared page count.",
    "A source worksheet and range are required.",
    "A structural operation cannot delete and move at the same time.",
    "A structural permutation requires an explicit move operation.",
    "A structural sort requires the host to authorize all worksheets so cross-sheet reference changes can be reviewed.",
    "A table already uses this name.",
    "A table cannot overlap merged cells. Unmerge the source first.",
    "A table header, totals row or calculated column cannot be the Goal Seek variable.",
    "A table with that name already exists.",
    "A temporary duplicate inspection context could not be isolated.",
    "A temporary watch inspection context could not be isolated.",
    "A theme needs eight colors and body and heading fonts.",
    "A transaction cannot change the same cell twice.",
    "A trial produced an array/spill; scalar results are required.",
    "A two-variable table needs two different inputs and exactly one formula result.",
    "A very-hidden worksheet must also be hidden.",
    "A what-if cell needs an Excel address.",
    "A what-if source is missing or outside the visible, authorized host context.",
    "A workbook may contain at most 1,000 defined names.",
    "A workbook may contain at most 100 formula watches, including orphaned entries. Remove a watch before adding another.",
    "A workbook may contain at most 100 formula watches.",
    "A workbook must contain at least one sheet.",
    "A workbook must have at least one sheet.",
    "A workbook supports at most 50 named styles.",
    "A workbook-level name has a context-dependent unqualified reference. Qualify its worksheet before moving cells.",
    "A worksheet references a missing cell style.",
    "A worksheet references a missing shared string.",
    "A worksheet supports at most 1,000 outline groups.",
    "A worksheet supports at most 100 drawing objects.",
    "A worksheet with that name already exists.",
    "A worksheet with that output name already exists.",
    "A worksheet-qualified expression is not a supported reference for a safe move.",
    "AMBIGUOUS",
    "Accrual exceeds 1000 quasi-coupon periods.",
    "Acknowledge conversion losses and compare the source before importing the converted workbook.",
    "Acknowledge that breaking links keeps the previewed cached values and removes the complete formulas.",
    "Acknowledge the retained parse exceptions before applying.",
    "Action",
    "Active formula watches must not repeat the same cell.",
    "Actor",
    "Add a border line before assigning its theme color.",
    "Add at least one transformation step before previewing.",
    "Adding this link would exceed the workbook populated-cell limit.",
    "Additional picture omitted: pictures in one worksheet may use at most 10 MB.",
    "Aggregation exceeds finite numeric bounds.",
    "Aging requires a Date, Due date, Invoice date, or Posting date column.",
    "Aging requires a numeric amount column.",
    "Aging requires a valid as-of date in YYYY-MM-DD format.",
    "All amounts are zero; no rate can be determined.",
    "Allow the print preview window in your browser, then try again.",
    "Alternate workbook XML declarations are unsupported.",
    "Alternate workbook XML exceeds its inspection budget.",
    "Alternate workbook XML must use UTF-8.",
    "Alternate workbook archive checksum mismatch.",
    "Alternate workbook archive directory mismatch.",
    "Alternate workbook archive exceeds its expansion budget.",
    "Alternate workbook archive is incomplete.",
    "Ambiguous alternate workbook identity.",
    "Ambiguous external workbook target relationship.",
    "Ambiguous source tables require review.",
    "Amount must be a finite number of at most 1e15 magnitude, exactly representable at the declared decimal precision; numeric text is not coerced.",
    "An XLSX numeric token exceeds 1,024 characters.",
    "An affected worksheet is protected. Unprotect it before changing dependent table formulas.",
    "An affected worksheet is protected. Unprotect it before changing names.",
    "An analysis label exceeds the 32,767-character cell limit. Shorten source labels or value-field names.",
    "An approved adapter with the exact extension version and manifest is unavailable.",
    "An empty pivot has no columns. Select at least one grouping key.",
    "An example row is outside the selected data rows.",
    "An extension cannot shadow a built-in function.",
    "An extension destination is outside the host write scope.",
    "An extension input attempted an unauthorized dependency.",
    "An extension input contains an unavailable calculation.",
    "An extension input is erroneous, stale, unauthorized or over its calculation limit.",
    "An extension input worksheet is unavailable or unauthorized.",
    "An image has one source page.",
    "An incomplete structured reference has no reliable formula signature.",
    "An inference example is outside the data rows or duplicated.",
    "An inferred input has an earlier parse exception.",
    "An internal hyperlink must target an existing worksheet.",
    "An internal hyperlink targets the moving totals row. Retarget or remove it before appending.",
    "An older workbook revision cannot overwrite a newer saved revision. Restore a checkpoint as a new revision instead.",
    "An operational calculation limit or cycle prevents screening, even when a formula handles the visible error.",
    "An original field source is outside the recorded range.",
    "An original filename cannot be inferred for a legacy picture.",
    "An outline group already exists for these rows or columns.",
    "An unbalanced structured reference prevents a safe move.",
    "An unknown or external worksheet qualifier cannot be safely migrated.",
    "An unresolved defined-name reference is present.",
    "An unsafe custom validation formula cannot be exported. Keep the native workbook and review the rule.",
    "An unsupported quoted reference cannot be safely migrated.",
    "An unterminated formula string cannot be safely migrated.",
    "An unterminated text literal has no reliable formula signature.",
    "Analysis and model report snapshots cannot be analysis inputs. Use the original source cells.",
    "Analysis definitions must be a bounded list.",
    "Analysis formulas require automatic, non-iterative calculation.",
    "Analysis metadata exceeds 2,000,000 characters.",
    "Analysis output exceeds 20,000 cells.",
    "Analysis permissions must be boolean.",
    "Analysis report exceeds the output budget.",
    "Analysis report metadata is inconsistent.",
    "Analysis result does not match its request.",
    "Another workbook window saved a different version. Reload or export your unsaved changes before resolving the conflict.",
    "Answer citations exceed the 1,000 cell-reference budget.",
    "Answer citations must stay inside the approved selection.",
    "Answer claim IDs must be unique.",
    "Answer claims exceed the 20,000 character evidence limit.",
    "Answer source is stale or outside the original approved context. Request a fresh answer.",
    "Append 1\u2013100 rows at a time.",
    "Append aligns headers; key normalization and fuzzy matching apply only to joins.",
    "Append matches headers; it does not accept join keys.",
    "Append requires exactly matching header names on both sources; column order may differ.",
    "Append requires the same exact header set; column order may differ.",
    "Approve distinct source qualifiers, at most eight workbooks.",
    "Approve one rectangular scope of at most 500 cells.",
    "Approve the host extractor before sharing the selected source.",
    "Approved selection changed. Review the proposal again.",
    "Approved workbook identities and worksheet permissions must be explicit, distinct and valid (at most 8 workbooks).",
    "Archive entries do not match its directory.",
    "Arithmetic paste requires All or Values mode; its results are numeric values.",
    "Array checking supports at most 1,000 formulas and 100,000 formula characters on the participating worksheets.",
    "Array discovery cannot infer blank or possible spill-child dependencies. Use a reviewed values snapshot before splitting.",
    "Array discovery encountered an operational calculation error. Resolve the uncertain formula state before splitting text.",
    "Array discovery exceeds 2,000 dependency cells. Use a smaller reviewed values snapshot.",
    "Array discovery requires a complete, authorized static dependency closure. Resolve hidden, dynamic, cyclic or unsupported formula dependencies before splitting.",
    "Array inspection supports at most 10,000 formulas and 10 million formula characters. Simplify the workbook before moving cells.",
    "Array-spilling formula migration is unsupported. Use explicit scalar formulas so every write is reviewable.",
    "Assignment requires a literal range.",
    "At least one positive and one negative cash flow are required.",
    "At least one row and one column must remain visible.",
    "At least one sheet must be visible.",
    "At least one worksheet must remain visible.",
    "At most 100 conditional formatting rules are supported.",
    "At most 100 validation rules are supported.",
    "At most 128 combined filter conditions are allowed.",
    "At most 40 saved analyses are supported. Remove an unused definition first.",
    "At most1,000 merged ranges are supported.",
    "Audit range exceeds sheet dimensions.",
    "AutoSum needs adjacent numbers or formulas above or to the left.",
    "Automation archive directory mismatch.",
    "Automation archive exceeds its expansion budget.",
    "Automation archive is incomplete.",
    "Automation artifact checksum does not match its preserved bytes.",
    "Automation artifact is not canonical base64.",
    "Automation artifact was not found.",
    "Automation artifacts exceed their 2 MB per-file or 10 MB retained-byte budget.",
    "Automation files must be 10 MB or smaller.",
    "Automation files must contain data and be 10 MB or smaller.",
    "Automation inspection was modified or is untrusted. Inspect the original bytes again.",
    "Automation inventory XML exceeds its bounds.",
    "Automation inventory does not accept XML document types.",
    "Automation run evidence is full (100 runs). Export native JSON before continuing.",
    "Automation run evidence is limited to 100 runs.",
    "Automation run revisions are inconsistent.",
    "Automation source must be UTF-8 text without null characters. Export a UTF-8 source copy; binary VBA is inventory-only.",
    "Automation source-line evidence must match the command count.",
    "Axis capacity must be a nonnegative integer.",
    "BLANK_KEY",
    "Binding reassignment is unsupported.",
    "Binomial points must be integers from zero through the explicit trial count; provide p in [0,1].",
    "Blank selection scans at most 100,000 cells. Select a smaller rectangle.",
    "Blank source cells require a bounded spill-ambiguity check; this sheet exceeds 1,000 possible formula anchors.",
    "Boolean values are not numeric fields.",
    "Boolean values must be TRUE or FALSE (case-insensitive).",
    "Border color roles must be an object.",
    "Both saved workbook snapshots are invalid. Original storage has been preserved for recovery.",
    "Browser storage is unavailable. Keep the workbook open and export your unsaved changes.",
    "Buckets must cover every nonnegative day exactly once: start at 0, use contiguous inclusive integer bounds, and leave only the final upper bound open.",
    "CARDINALITY",
    "CSV/TSV export exceeds the 10 MB output limit. Choose a smaller range or native JSON.",
    "CYCLE",
    "Calculated column formulas must begin with = and contain at most 4,096 characters, without outer whitespace.",
    "Calculated comparison cannot group formula errors or error-like formula strings; correct them or use raw comparison.",
    "Calculated comparison excludes volatile, dynamic-reference and array/spill formulas. Raw comparison can preserve their formula text without execution.",
    "Calculated comparison requires a complete supported static dependency closure within 1,000 cells, 5,000 scanned formulas and depth 8; hidden, dynamic, unsupported or cyclic dependencies are excluded.",
    "Calculated comparison requires populated source cells; blank or possible spill-child cells are outside this subset. Use raw comparison or a populated selection.",
    "Calculated comparison requires populated static dependencies; blank or possible spill-child dependencies are unsupported.",
    "Calculated comparison requires scalar cells; a dependency produced an array/spill result.",
    "Calculated field names must be unique.",
    "Calculated searches support at most 10,000 cells. Narrow the selection or use Raw contents.",
    "Calculated sort keys changed after the preview. Preview and review the new order.",
    "Calculated source coverage is incomplete or encountered an operational error.",
    "Calculated subtotals do not reconcile with the reviewed amounts, or a calculation limit was reached. Reduce numeric cancellation or explicitly round source amounts before publishing.",
    "Calculated table formulas require explicit table append/resize. Clear their saved rules (keeping cell formulas) before inserting or deleting worksheet rows/columns, then reapply the rules.",
    "Calculated tables support downward resize only; keep the existing header and columns. Clear rules before shrinking or changing columns.",
    "Calculated values or population membership changed. Preview the export again.",
    "Calculation cancelled. Inputs are intact; use Calculate Now to retry.",
    "Calculation clock is outside the supported 1900\u20139999 date range.",
    "Calculation exceeded its time budget and its worker was terminated. Inputs are intact; simplify formulas or increase the budget, then retry.",
    "Calculation inputs exceed the working-data budget. Reduce the data or increase the budget in Calculation Options.",
    "Calculation options support only filteredRowsBySheet.",
    "Calculation results exceed the working-data budget; no partial results were published.",
    "Calendar and weekday series require a real calendar seed; use Days for the fictitious 1900-02-29 serial.",
    "Cash-flow input exceeds 10000 positions.",
    "Cash-flow signs or payment size do not produce a finite real period count.",
    "Cell borders support top, right, bottom and left edges only.",
    "Cell entry must be text.",
    "Cell is outside this citation.",
    "Cell literal flag must be boolean.",
    "Cell note",
    "Cell style must be an object.",
    "Cell value",
    "Cell values must be text, finite numbers, booleans, or null.",
    "Change history has reached its 100,000 cell-change limit. Export the workbook before starting a new workpaper.",
    "Change history has reached its 100,000 cell-change limit. Export this workbook before continuing.",
    "Change history has reached its 2,000 transaction limit. Export the workbook before starting a new workpaper.",
    "Change history is full. Export this workbook before continuing.",
    "Changing cell must contain a finite numeric input, not a formula, text, blank or spill output.",
    "Changing cells must be unique.",
    "Changing cells must contain finite numeric inputs, not formulas or text.",
    "Changing cells must contain finite numeric inputs, not formulas, text, blanks or spill outputs.",
    "Changing cells with validation rules are excluded. Use explicit Solver bounds and constraints on a reviewed input copy.",
    "Changing inputs and formula results must be different cells.",
    "Chart filters require at most 500 unique one-based worksheet rows.",
    "Chart inspection supports at most 20 value columns; narrow the source range.",
    "Chart scales must be linear or log.",
    "Chart series names require 1 to 200 characters.",
    "Chart series require unique one-based worksheet columns.",
    "Checkpoint label must contain 1\u2013120 characters.",
    "Choose 1 to 50 histogram bins.",
    "Choose 1\u201312 text cleanup steps.",
    "Choose 1\u201320 explicit worksheet/range mappings.",
    "Choose 1\u20138 available fields for inference.",
    "Choose 1\u20138 delimiter characters without quotes, line breaks or NUL.",
    "Choose Body or Heading theme font.",
    "Choose Copy or Series for the fill handle.",
    "Choose JSON, XML, PDF, PNG, JPEG or HTML.",
    "Choose Number, Date or Text for ordered comparisons.",
    "Choose Save model or Publish result.",
    "Choose Save report or Apply solution and report.",
    "Choose Top, Center or Bottom vertical alignment.",
    "Choose VBA .bas/.vba/.cls, Office Script .ts/.osts, or an Excel container for inventory.",
    "Choose a JSON object or a nonempty array.",
    "Choose a PNG or JPEG picture no larger than 2 MB.",
    "Choose a PivotTable value cell, subtotal or grand total; headers and group labels cannot be drilled.",
    "Choose a configuration no larger than 2,000,000 bytes.",
    "Choose a configuration/report no larger than 2,000,000 bytes.",
    "Choose a conversion mode, exception resolution and different decimal/group separators.",
    "Choose a different source qualifier.",
    "Choose a distinct worksheet name of 1\u201331 characters without invalid characters.",
    "Choose a document source for extraction.",
    "Choose a formula worksheet and variable worksheet.",
    "Choose a fuzzy key pair and a similarity threshold from 0 to 1.",
    "Choose a header range and up to 64 column filters.",
    "Choose a hexadecimal color.",
    "Choose a horizontal split, vertical split, or both.",
    "Choose a native JSON or XLSX comparison baseline.",
    "Choose a native workbook JSON or an XLSX file.",
    "Choose a navigation direction.",
    "Choose a new output worksheet name.",
    "Choose a numeric or date series.",
    "Choose a pattern-audit configuration or report.",
    "Choose a picture first.",
    "Choose a saved recipe ID and version together.",
    "Choose a saved recipe version.",
    "Choose a separate destination so the original source cells remain intact.",
    "Choose a separate output range or a new worksheet.",
    "Choose a six-digit hexadecimal color.",
    "Choose a source range with at least one data row and valid header handling.",
    "Choose a source range, cleanup steps and output worksheet name.",
    "Choose a source worksheet and one rectangular range.",
    "Choose a source worksheet and one-column range.",
    "Choose a source worksheet.",
    "Choose a supported Go To Special category.",
    "Choose a supported Look in option.",
    "Choose a supported Show values as setting.",
    "Choose a supported aging configuration or report.",
    "Choose a supported border color role.",
    "Choose a supported border line style and hexadecimal color.",
    "Choose a supported border preset.",
    "Choose a supported chart type.",
    "Choose a supported date unit and month policy.",
    "Choose a supported fill direction.",
    "Choose a supported paste mode and arithmetic operation.",
    "Choose a supported screen configuration or report.",
    "Choose a supported table column type.",
    "Choose a supported theme color role.",
    "Choose a supported theme font.",
    "Choose a supported theme or named style action.",
    "Choose a supported workbook interface locale.",
    "Choose a table operation.",
    "Choose a target workbook link.",
    "Choose a unique worksheet name of 1\u201331 characters.",
    "Choose a valid fill first.",
    "Choose a valid fixed as-of date in ISO YYYY-MM-DD format.",
    "Choose a valid output worksheet name of 1\u201331 characters.",
    "Choose a valid pivot field.",
    "Choose a valid position and 1\u201310,000 rows or columns.",
    "Choose a valid row or column change for review targets.",
    "Choose a valid search scope and order.",
    "Choose a valid source, 1\u201310,000 rows or columns, and a destination within the worksheet.",
    "Choose a valid structural edit of 1\u201310,000 rows or columns.",
    "Choose a valid visible-axis slice.",
    "Choose a valid worksheet cell address.",
    "Choose a web, email or worksheet link.",
    "Choose a workbook theme.",
    "Choose a worksheet and one explicit headed cell rectangle for each source.",
    "Choose a worksheet position inside the workbook.",
    "Choose an XLS, XLSB, XLSM, XLTX, XLTM or ODS source.",
    "Choose an XLSX, CSV, TSV or JSON file.",
    "Choose an absolute or signed amount and a supported comparison rule.",
    "Choose an aging configuration/report no larger than 2,000,000 bytes.",
    "Choose an available outline level.",
    "Choose an existing output worksheet.",
    "Choose an existing source worksheet.",
    "Choose an output outside tables, merged areas and generated analysis results.",
    "Choose an output worksheet and one starting cell.",
    "Choose an unpadded output worksheet name of 1\u201331 characters.",
    "Choose append or a supported join operation.",
    "Choose at least one output field.",
    "Choose at least two data rows. Check the header setting.",
    "Choose between 1 and 100 pivot value fields.",
    "Choose distinct field positions from 1 to 50.",
    "Choose distinct fill group columns.",
    "Choose distinct valid sort columns, directions and data types.",
    "Choose due or document date and an explicit 1900 or 1904 serial date system.",
    "Choose explicit worksheet and rectangular range references.",
    "Choose from one to eight pairs of join headers.",
    "Choose full or visible analysis scope.",
    "Choose headers and 1\u201364 sort levels with valid options.",
    "Choose next or previous.",
    "Choose one cell to inspect.",
    "Choose one date grouping per source field.",
    "Choose one delimiter character other than a quote, line break or NUL.",
    "Choose one destination cell or a rectangle with exactly the source dimensions.",
    "Choose one explicit report range.",
    "Choose one or more columns inside the selected range.",
    "Choose one or more key columns inside the source range.",
    "Choose one rectangular A1 source range.",
    "Choose one reviewed scenario.",
    "Choose only one control of each kind per field.",
    "Choose precedents or dependents.",
    "Choose rows or columns.",
    "Choose scalar fields inside one source rectangle of at most 500 cells and within worksheet dimensions.",
    "Choose scenarios, a one-variable table or a two-variable table.",
    "Choose the converted XLSX copy; renaming the original does not convert it.",
    "Choose the merged area's first cell or unmerge it before changing its link.",
    "Choose three distinct amount, currency and date columns inside the source range.",
    "Choose two increasing icon thresholds; percentages must be between 0 and 100.",
    "Choose valid ISO start and end dates in chronological order.",
    "Choose valid comparison, header, case, trim and blank rules.",
    "Choose valid worksheet size settings.",
    "Choose visible, hidden or very hidden.",
    "Choose whether the first source row contains headings.",
    "Choose whether to resolve or reopen the discussion.",
    "Citation includes hidden cells. Make the cited evidence visible and request a fresh answer.",
    "Citation is outside approved context.",
    "Citation range must run top left to bottom right.",
    "Citation requires an explicit A1 cell or rectangle.",
    "Citation worksheet is unavailable or hidden.",
    "Citations require one approved rectangle.",
    "Citations support at most 500 cells per range.",
    "Cleanup output exceeds the 100,000 stored-cell budget.",
    "Clear calculated column rules before changing table columns.",
    "Clear calculated column rules before changing totals-row visibility.",
    "Clear saved calculated-column rules before moving worksheet structures, keeping their cell formulas; then reapply the rules.",
    "ClearContents requires a range.",
    "Clipboard cells must form a complete row-major rectangle.",
    "Clipboard dimensions do not match its source range.",
    "Clipboard text contains malformed quoted cells.",
    "Clipboard text is too large.",
    "Code outside the single Sub is unsupported.",
    "Collapsed state must be boolean.",
    "Column",
    "Column conversion must be Text or General, with at most 50 columns.",
    "Column count",
    "Column index is outside the worksheet.",
    "Column names cannot contain brackets or line breaks.",
    "Column widths must be 40\u2013600 pixels for an existing column.",
    "Column widths must be an index map.",
    "Combined source formula dependencies exceed 2,000 cells.",
    "Combined source text exceeds 1,000,000 characters.",
    "Combo series must be column or line.",
    "Comment",
    "Comparison exceeds 30,000 findings; narrow the source workbooks.",
    "Comparison mappings may not overlap on either side.",
    "Comparison supports at most 10,000 selected cells across both workbooks.",
    "Comparisons require two numbers or two text values of the same type.",
    "Complete answer evidence exceeds 2,000,000 bytes. Reduce the selection or answer; evidence is never silently truncated.",
    "Concatenation would exceed 50 fields.",
    "Conditional format preview is stale. Reopen the manager to refresh it.",
    "Conditional format preview was modified. Preview the rules again.",
    "Conditional formula evaluation exceeds 5,000 cells; narrow the visible or preview range.",
    "Conditional priorities must be unique integers from 1 to 100.",
    "Conditional rule IDs must be nonempty and unique.",
    "Configure between 1 and 24 aging buckets.",
    "Confirm the deletion and reference changes before applying.",
    "Consolidation evidence does not reconcile to its sources, categories or retained fingerprint.",
    "Consolidation evidence exceeds 8,000,000 characters. Retain native backups before removing obsolete reports.",
    "Consolidation exceeds 1,000 subtotal groups. Narrow the sources; no rows were truncated.",
    "Consolidation exceeds 2,000 contributions. Narrow the sources; no output was truncated.",
    "Consolidation exceeds the output or workbook cell budget. No worksheet was published.",
    "Consolidation output range exceeds its cell budget.",
    "Consolidation source evidence does not reconcile.",
    "Consolidation total exceeds finite numeric limits.",
    "Conversion would exceed safe numeric precision. Keep the original text.",
    "Copy at most 500 cells into the internal clipboard.",
    "Correlation is undefined for a constant column.",
    "Corrupt XLSX compressed content. Repair the original workbook and save a clean XLSX copy.",
    "Corrupt alternate workbook archive.",
    "Cost and salvage must be nonnegative.",
    "Could not capture spill results.",
    "Count rows does not take a column.",
    "Coupon frequency must be 1, 2 or 4.",
    "Create a preview first.",
    "Creating a derived worksheet is not permitted.",
    "Creating an output worksheet is not permitted while workbook structure is protected or the host disallows new worksheets.",
    "Creating an output worksheet is not permitted.",
    "Custom function arguments do not match the declared signature.",
    "Custom orders require Text, 1\u2013100 distinct nonempty labels of at most 255 characters.",
    "Cut and move supports at most 500 source cells.",
    "DAX measures require the separate model expression evaluator.",
    "DAX_ENGINE",
    "DB requires whole life/period and 1\u201312 first-year months; a partial first year permits one partial final year.",
    "DDB factor must be greater than zero.",
    "DDE-style syntax is inventoried and never executed.",
    "DISCONNECTED",
    "DTD is not permitted in preservation metadata.",
    "DUPLICATE_KEY",
    "Data model execution exceeded 2,000,000 operations. Reduce the model or query.",
    "Data model metadata exceeds 2,000,000 characters.",
    "Data validation exceeds the 100,000-check budget; split the batch.",
    "Date grouping requires a year, quarter or month and an explicit ISO/1900/1904 date system.",
    "Date must be strict ISO YYYY-MM-DD or a valid Excel date/time serial in the chosen system; serial 60 in the 1900 system is invalid.",
    "Date must match the selected year/month/day order, optionally HH:mm:ss.",
    "Dated cash flows require one numeric amount for every date; blanks, booleans and text cannot be silently removed.",
    "Dates require a four-digit year.",
    "Day-count basis must be 0 (US 30/360), 1 (actual/actual), 2 (actual/360), 3 (actual/365), or 4 (European 30/360).",
    "Decimal and grouping separators must differ.",
    "Declare a three-letter currency identifier; no currency conversion is performed.",
    "Decoded picture dimensions do not match its validated header.",
    "Default-hidden worksheet rows cannot be imported safely. Replace the default with explicit hidden rows before importing.",
    "Defined names contain a circular reference.",
    "Defined-name comments may contain at most 1,000 characters.",
    "Defined-name cycle: references cannot be completely inspected.",
    "Defined-name expansion exceeds depth 16.",
    "Defined-name worksheet qualifier is missing.",
    "Defined-name worksheet scope was not found.",
    "Deletion is blocked by ambiguous or unsupported references. Replace dynamic or three-dimensional references with explicit references first.",
    "Delimited data exceeds 100,000 populated cells.",
    "Delimited export exceeds the 10 MB output limit. Choose a smaller range.",
    "Delimited files cannot contain null bytes; check the encoding.",
    "Delimited files must be 10 MB or smaller.",
    "Delimited source footprint exceeds worksheet bounds or has invalid header metadata.",
    "Delimited/CSV export would expand more than 100,000 cells including blank gaps. Choose a smaller range or use native JSON.",
    "Dependency inspection encountered an operational calculation error.",
    "Depreciation period must be a supported whole period within asset life (plus DB partial final year).",
    "Derived calculation exceeds finite numeric bounds.",
    "Derived expression divides by zero.",
    "Derived output exceeds the 100,000 stored-cell workbook budget.",
    "Derived text exceeds the cell text limit.",
    "Derived value exceeds finite numeric or text limits.",
    "Diagnostic history must retain between 2 and 1000 events.",
    "Distribution primitives cannot read workbook cells.",
    "Distribution primitives cannot read workbook ranges.",
    "Document cells require a source page and bounding region.",
    "Document extraction cancelled.",
    "Document extraction exceeded its time budget.",
    "Document extraction failed. The workbook is unchanged.",
    "Drawing IDs must be unique.",
    "Drawing anchor exceeds the worksheet.",
    "Drawing colors must be six-digit hex colors.",
    "Drawing footprint exceeds the move inspection budget.",
    "Drawing print inspection exceeds its geometry budget.",
    "Drill context exceeds the cell-note budget. Reduce filter selections before extracting.",
    "Drill row associations exceed the 100,000 source-row budget.",
    "Duplicate XLSX calculation properties. Repair the source workbook.",
    "Duplicate XLSX page-break sections are not supported.",
    "Duplicate automation artifact ID.",
    "Duplicate automation run ID.",
    "Duplicate citation ranges in a claim.",
    "Duplicate custom function names.",
    "Duplicate extension cache entries.",
    "Duplicate extension parameter names.",
    "Duplicate headers are ambiguous; resolve them or compare without headers.",
    "Duplicate key is blank or has an invalid value.",
    "Duplicate or late VBA declaration.",
    "Duplicate or missing alternate workbook content type.",
    "Duplicate or over-capacity field transformation evidence.",
    "Duplicate original field source.",
    "Duplicate report output IDs.",
    "Duplicate source import ID.",
    "Duplicate transformed output address.",
    "Duplicate/unique conditional export exceeds 1,000,000 population checks. Use native JSON or fewer rules.",
    "Duplicating this worksheet would exceed the 100,000 populated-cell budget.",
    "Duplicating worksheets with three-dimensional references requires explicit references first.",
    "Dynamic reference function is not executed; actual target is unverified.",
    "Dynamic spill references require an explicit stable range before moving cells.",
    "Each advanced OR group must contain at least one AND condition.",
    "Each appended row must contain one raw value for every table column.",
    "Each bucket needs a unique label of 1\u201380 characters and inclusive from/to bounds.",
    "Each cell must have an explicit address inside its source row.",
    "Each changing input must be a precedent of at least one selected formula result.",
    "Each column supports up to two custom conditions.",
    "Each comparison range supports at most 1,000 rows and 100 columns within the worksheet.",
    "Each mapping requires explicit header and row-matching rules.",
    "Each scenario needs exactly one value per changing input.",
    "Each source supports at most 500 data rows, 50 columns and 2,000 cells including headers.",
    "Embedded image exceeds 2 MB.",
    "Enable chat support on the host proxy before sending conversation data.",
    "Encrypted Office or legacy binary workbooks are unsupported. Remove the open password in the source application.",
    "Encrypted OpenDocument workbooks require an unencrypted XLSX conversion.",
    "Encrypted XLSX archive entries are unsupported.",
    "Engagement must contain at most 200 characters.",
    "Enter 1\u2013256 characters to find. Use Go To Special to select blank cells.",
    "Enter TRUE or FALSE for Boolean fill.",
    "Enter a column letter such as A or XFD, or use end for the move destination.",
    "Enter a complete https://, http:// or mailto: address.",
    "Enter a date as YYYY-MM-DD or a whole Excel date serial.",
    "Enter a finite fill number.",
    "Enter a finite number.",
    "Enter a finite target value.",
    "Enter a print area.",
    "Enter a step value.",
    "Enter a valid output worksheet name of at most 31 characters.",
    "Enter an address, range, or defined range name (up to 8,192 characters).",
    "Enter an existing row number, or use end for the move destination.",
    "Enter an output worksheet name.",
    "Enter at least two worksheet row numbers, separated by commas.",
    "Enter cell addresses separated by commas or new lines.",
    "Enter comma-separated row numbers or column letters for manual breaks.",
    "Enter field positions separated by commas, such as 1, 3.",
    "Enter finite comparison values.",
    "Enter finite validation bounds.",
    "Enter increasing fixed-width break positions separated by commas.",
    "Enter text to find.",
    "Enter valid comma-separated addresses or defined range names.",
    "Entity results require the linked-data permission.",
    "Error selection calculates at most 10,000 cells. Narrow the selection.",
    "Event ID",
    "Event timestamp",
    "Every cash-flow amount must have exactly one date.",
    "Every contributing row needs nonblank categories.",
    "Every join key must name an existing source header exactly.",
    "Every result cell must contain an actual calculation formula.",
    "Every rule needs a unique ID.",
    "Every sort key must be inside the complete record range.",
    "Every source column needs a nonblank text header of at most 128 characters.",
    "Every source field already has a slicer.",
    "Every what-if input value must be a finite number.",
    "Evidence metadata belongs only to pictures.",
    "Excel number format",
    "Executable or embedded-object relationships are unsupported.",
    "Expanded Solver expression budget exceeded.",
    "Expanded output text exceeds 1,000,000 characters.",
    "Expected TRUE or FALSE.",
    "Expected a finite number.",
    "Expected a finite number; nonnumeric text cannot be used as a financial input.",
    "Expected a full-year ISO date or an explicitly ordered slash date.",
    "Expected a safe whole number.",
    "Explicit approval is required before sending extension arguments.",
    "Explicit approval of the migrated commands is required.",
    "Explicit exclusions must be distinct data-row numbers inside the selected range.",
    "Explicitly approve activation of inert imported formulas before resolving them.",
    "Exponential smoothing needs at least three observations.",
    "Exporting print identity fields requires the current workbook snapshot.",
    "Exporting this merged area would discard populated cells or annotations. Unmerge it or preserve those values before XLSX export.",
    "Extend the fill handle along one axis while keeping the other source dimension unchanged.",
    "Extension argument or call count exceeds its review limit.",
    "Extension arguments cannot be omitted.",
    "Extension cache does not match an installed function.",
    "Extension cache identity does not match its arguments.",
    "Extension functions cannot shadow built-ins or defined names.",
    "Extension functions cannot shadow defined names.",
    "Extension identities/namespaces must be unique within the function limit.",
    "Extension matrix exceeds its declared shape or cell limit.",
    "Extension matrix must be rectangular.",
    "Extension metadata exceeds its limit.",
    "Extension processing cancelled or timed out.",
    "Extension processing cancelled.",
    "Extension results exceed the total output cell budget.",
    "Extension review exceeds 100 formula cells.",
    "Extension timeout must be between 1 and 60,000 milliseconds.",
    "Extension value does not match its declared schema.",
    "Extensions must declare distinct read-arguments and return-results permissions.",
    "External claims require distinct references to declared external sources.",
    "External link metadata cannot contain DTDs.",
    "External link refresh exceeds 5,000 source cells.",
    "External links require bounded A1 cells or rectangles.",
    "External or executable custom functions are not supported.",
    "External or three-dimensional references cannot be safely migrated.",
    "External or three-dimensional references prevent a safe range move.",
    "External or unqualified bracket references cannot be safely migrated.",
    "External reference exceeds worksheet bounds.",
    "External runtimes must declare remote-compute.",
    "External source IDs must be unique.",
    "External sources require a valid HTTPS URL without credentials.",
    "External workbook link is inventoried but never opened; target validity is unverified.",
    "External workbook references are not opened or fetched.",
    "External workbook references prevent a safe move.",
    "External-action conditional formulas cannot be exported. Remove that rule or use native JSON.",
    "External/network-enabled function is inventoried but never executed.",
    "Extraction does not match the selected source hash.",
    "Extraction has invalid types, confidence, regions or source identity.",
    "FILTER",
    "Facet column is outside the filter range.",
    "Field source evidence must contain reviewed values.",
    "Field transformation evidence does not match its recipe.",
    "Field transformation evidence exceeds its native history budget. Keep the original backup.",
    "File worker unavailable.",
    "Files must be 10 MB or smaller.",
    "Fill at most 500 destination cells.",
    "Fill targets must differ from the group columns. Use a separate preceding step to prepare group keys.",
    "Filter columns must be valid and distinct.",
    "Filter comparison work exceeds its 5,000,000-unit budget. Narrow the range or simplify criteria.",
    "Filter context refers to an unknown worksheet ID.",
    "Filter context supports at most 10,000 intervals across the workbook.",
    "Filter criteria text is limited to 1,024 characters.",
    "Filter range needs one header row and at least one data row within the worksheet.",
    "Filter values must be finite numbers, text, booleans or blank.",
    "Filter values must match the declared column type.",
    "Filtered exports require a header and an explicit full or visible scope.",
    "Filtered rows must be a worksheet-ID map of inclusive row intervals.",
    "Filtering is limited to 10,000 data rows, 100,000 cells including headers, and 64 columns. Choose a smaller range.",
    "Financial periods exceed the 10000-period budget.",
    "Finish each corrected value before previewing.",
    "Finish the current cell edit and request fresh evidence before following a citation.",
    "Fixed-width breaks must be distinct increasing character positions.",
    "Font size must be an integer between 8 and 72.",
    "Forecast periods must be strictly increasing and equally spaced without gaps.",
    "Forecasting requires two columns: increasing equally spaced numeric period, then value.",
    "Formula commands require explicit A1 formulas beginning with =.",
    "Formula dependencies must be populated scalar cells; ambiguous blank or spill-child dependencies cannot be screened.",
    "Formula dependencies must be populated static cells; blank or possible spill-child dependencies are outside this subset.",
    "Formula exceeds the 4,096-character inspection bound.",
    "Formula fields require a complete supported static dependency closure within the visible authorized context (1,000 cells, depth 8, 5,000 scanned formulas).",
    "Formula must return a scalar boolean or number.",
    "Formula offsets must be integers.",
    "Formula watch IDs must be unique.",
    "Formula-like value assignment has different Excel semantics. Use an explicit setFormula / .Formula statement or change the reviewed source.",
    "Freeze panes cannot cross a merged range. Unmerge those cells first.",
    "Freeze panes must leave an unfrozen visible row and column.",
    "Freeze panes supports at most 100 visible rows and 50 visible columns.",
    "Frozen boundaries must be valid leading row and column counts.",
    "Fuzzy comparison exceeds 10,000,000 work units. Narrow the populations or add an exact key before previewing.",
    "Fuzzy keys are limited to 512 Unicode characters. Shorten the reviewed key or use exact matching.",
    "Generated pivot/query results cannot be what-if inputs.",
    "Generated results cannot be changing cells.",
    "Go To Special may select at most 500 cells. Narrow the selection.",
    "Go To Special scans at most 100,000 cells.",
    "Go To Special selects cells on one sheet. Choose Selection or Sheet.",
    "Goal Seek does not run volatile, dynamic-reference, array/spill or unsupported reusable-function dependencies.",
    "Goal Seek host scope contains an unknown worksheet.",
    "Goal Seek needs one valid Excel cell address for each source.",
    "Goal Seek requires a complete dependency closure within 500 cells, 5,000 scanned formulas and depth 8.",
    "Goal Seek requires populated static dependency cells; blank or possible spill-child dependencies are not supported.",
    "Goal Seek supports continuous numeric/date validation only; discrete list, integer or text-length constraints need a different input.",
    "Goal Seek supports from 1 to 100 iterations.",
    "Grand totals must be boolean.",
    "Group and measure headings must be distinct.",
    "Grouped date exceeds year 9999.",
    "Grouping separators must separate groups of three digits.",
    "Hashing requires at most 50 MB of bytes.",
    "Header and whitespace controls must be boolean.",
    "Header rows require nonempty literal text in every column. Turn off headers for an unstructured sheet.",
    "Headers and footers must be at most 255 characters.",
    "Hidden XLSX row numbers must be distinct and ascending.",
    "Hidden XLSX rows must have valid row numbers.",
    "Hidden indices require at most 10,000 intervals per axis.",
    "Hidden intervals require valid zero-based inclusive axis bounds.",
    "Hidden metadata exceeds 10,000 intervals per worksheet axis.",
    "Hidden state must be boolean.",
    "Hidden worksheets are outside the automation scope.",
    "Histogram bin upper bounds must be finite and strictly increasing.",
    "History cell address is not canonical.",
    "History change must have a before or after value.",
    "History events require unique IDs and valid timestamps.",
    "Holiday dates must be distinct valid ISO calendar dates.",
    "Host read-only and hidden-source policies must be boolean.",
    "Host read-only must be boolean.",
    "Host selection changed while reading. Retry the request.",
    "IF/AND/OR/NOT conditions must be Boolean comparisons.",
    "Imported source cell is unavailable.",
    "Include all table columns when removing duplicate table rows.",
    "Include hidden must be boolean.",
    "Incomplete Solver formula.",
    "Indent must be an integer between 0 and 15.",
    "Inert imported formulas must be deliberately entered as active formulas before extension execution.",
    "Inputs or permissions changed during extension processing.",
    "Insertion would discard a review target or hyperlink at the worksheet boundary.",
    "Insertion would discard explicit sizing at the worksheet boundary.",
    "Insertion would discard hidden-axis metadata at the worksheet boundary.",
    "Insertion would discard outline groups at the worksheet boundary. Ungroup those rows or columns first.",
    "Insertion would discard source-row provenance at the worksheet boundary.",
    "Insertion would move existing content beyond the Excel worksheet boundary.",
    "Inspect at most 30 pivot control fields together.",
    "Inspection ranges must remain within authorized worksheets.",
    "Insufficient numeric observations after exclusions.",
    "Integer column requires a safe whole number.",
    "Interest/principal period must be an integer from 1 through nper.",
    "Intermediate results exceed the 1,000,000-character limit.",
    "Internal calculation names cannot be called from a workbook.",
    "Invalid Goal Seek request.",
    "Invalid JSON array.",
    "Invalid JSON field.",
    "Invalid JSON object.",
    "Invalid JSON string.",
    "Invalid JSON value.",
    "Invalid LET/LAMBDA bindings prevent complete dependency inspection.",
    "Invalid Solver function arguments.",
    "Invalid XLSX archive directory.",
    "Invalid XLSX archive offsets.",
    "Invalid XLSX calculation properties. Repair the source workbook.",
    "Invalid XLSX page-break boundary.",
    "Invalid XLSX streaming data descriptor.",
    "Invalid XLSX worksheet metadata.",
    "Invalid XML nil value.",
    "Invalid advanced filter groups.",
    "Invalid alternate workbook XML identity.",
    "Invalid alternate workbook filename or 10 MB file limit exceeded.",
    "Invalid alternate workbook main relationship.",
    "Invalid alternate workbook source or conversion provenance.",
    "Invalid analysis identity or source revision.",
    "Invalid analysis metadata.",
    "Invalid answer grounding: provide typed claims and separate external sources.",
    "Invalid approved workbook identity or worksheet scope.",
    "Invalid authorized worksheet context.",
    "Invalid authorized worksheet scope.",
    "Invalid automation approved range.",
    "Invalid automation artifact encoding or size.",
    "Invalid automation artifact.",
    "Invalid automation evidence count.",
    "Invalid automation file name.",
    "Invalid automation run evidence.",
    "Invalid automation source hash.",
    "Invalid automation source lines.",
    "Invalid automation timestamp.",
    "Invalid binary workbook beginning record. Convert the original in its source application.",
    "Invalid calculation settings. Choose Automatic or Manual, explicit iteration, 1\u20131000 iterations, convergence from 0 to 1, a 1\u201310000 ms time budget, 1\u2013128 MiB working data and 1000\u201310000000 work units.",
    "Invalid cash-flow or security date. Use a valid Gregorian/workbook-system serial or explicit full-year date text.",
    "Invalid cell patch.",
    "Invalid chart header setting.",
    "Invalid chart legend setting.",
    "Invalid column type request.",
    "Invalid column type selection.",
    "Invalid conditional rule.",
    "Invalid conditional-format stop flag.",
    "Invalid data validation rule.",
    "Invalid date export mode.",
    "Invalid date/time component.",
    "Invalid defined name.",
    "Invalid delimited source footprint.",
    "Invalid derived operation.",
    "Invalid destination worksheet scope.",
    "Invalid document extraction time budget.",
    "Invalid drawing zoom.",
    "Invalid duplicate review request.",
    "Invalid embedded-object inventory count or size.",
    "Invalid explicit key normalization options.",
    "Invalid extension host scope.",
    "Invalid external range.",
    "Invalid field evidence timestamp or incomplete original source population.",
    "Invalid filter condition, column or value.",
    "Invalid formula watch metadata.",
    "Invalid formula watch orphan reason.",
    "Invalid history cell reference.",
    "Invalid history change count.",
    "Invalid history event.",
    "Invalid host extension registration.",
    "Invalid icon threshold type.",
    "Invalid imported formula-cache evidence. Caches must remain separately labeled source evidence.",
    "Invalid isolated format inspection.",
    "Invalid isolated inspection.",
    "Invalid local workbook envelope.",
    "Invalid missing theme flag.",
    "Invalid named style fallback flag.",
    "Invalid named style.",
    "Invalid number for the selected decimal/grouping separators.",
    "Invalid or duplicate XLSX content type declaration.",
    "Invalid or duplicate conditional-format priority.",
    "Invalid or over-capacity source row association.",
    "Invalid or over-capacity table source evidence. Keep the native backup.",
    "Invalid or over-capacity text cleanup evidence. Keep the original native backup.",
    "Invalid or oversized analysis output range.",
    "Invalid or oversized internal clipboard snapshot.",
    "Invalid or unknown chart settings.",
    "Invalid or unsafe typed XML number.",
    "Invalid orphaned review target.",
    "Invalid output destination.",
    "Invalid output encoding or BOM setting.",
    "Invalid page layout.",
    "Invalid paper, orientation, or margins.",
    "Invalid picture evidence metadata.",
    "Invalid picture evidence origin.",
    "Invalid pivot aggregation. Choose sum, count, average, min, max or distinctCount.",
    "Invalid pivot content fingerprint.",
    "Invalid pivot definition.",
    "Invalid pivot measure count.",
    "Invalid pivot measure evidence.",
    "Invalid pivot numeric count.",
    "Invalid pivot population counts.",
    "Invalid pivot refresh evidence.",
    "Invalid pivot refresh time or source label.",
    "Invalid pivot slicer.",
    "Invalid print option.",
    "Invalid query definition.",
    "Invalid query sort.",
    "Invalid recipe row width.",
    "Invalid report cell address.",
    "Invalid report cell layout.",
    "Invalid report column layout.",
    "Invalid report drawing identity.",
    "Invalid report row accounting.",
    "Invalid retained XLSX source encoding.",
    "Invalid retained XLSX source metadata or 10 MiB source limit exceeded.",
    "Invalid retained source encoding.",
    "Invalid review resolution time.",
    "Invalid saved advanced criteria.",
    "Invalid saved cell range.",
    "Invalid saved filter columns.",
    "Invalid saved filter conditions.",
    "Invalid saved filter view identity, revision or timestamp.",
    "Invalid saved filter view metadata.",
    "Invalid saved population counts.",
    "Invalid saved what-if assumption.",
    "Invalid saved what-if result dimensions or outcome.",
    "Invalid saved worksheet selection.",
    "Invalid selection enumeration limit.",
    "Invalid sheet protection.",
    "Invalid sheet.",
    "Invalid slicer item.",
    "Invalid source filter.",
    "Invalid source review.",
    "Invalid sparkline style.",
    "Invalid stored XLSX entry size.",
    "Invalid string escape.",
    "Invalid structural edit bounds.",
    "Invalid table column.",
    "Invalid table metadata.",
    "Invalid table total field or operation.",
    "Invalid template provenance.",
    "Invalid text alignment.",
    "Invalid typed XML boolean.",
    "Invalid validation display setting.",
    "Invalid validation error style.",
    "Invalid what-if analysis fields.",
    "Invalid what-if evidence counts.",
    "Invalid workbook metadata fields.",
    "Invalid workbook property fields.",
    "Invalid workbook structure protection.",
    "Invalid worksheet drawing object.",
    "Invalid worksheet tool text.",
    "Invalid workspace",
    "Invalid, duplicated or out-of-bounds original source row.",
    "Isolated calculation is unavailable. The host must allow or provide a calculation worker; source inputs are intact.",
    "JPEG dimensions conflict.",
    "JPEG frame components are invalid.",
    "JPEG has too many segments.",
    "JPEG has trailing data.",
    "JPEG marker structure is invalid.",
    "JPEG scan is invalid.",
    "JPEG scan or end marker is missing.",
    "JPEG segment is truncated.",
    "JPEG segment length is invalid.",
    "JSON nesting or node budget exceeded.",
    "JSON/XML use the native structured parser.",
    "Keep at most 20 automation artifacts.",
    "Keep the table header row selected as headers.",
    "Keep your message under 10,000 characters.",
    "Key headers must exist exactly once and may not be repeated.",
    "Key matching requires 1\u20138 explicit header pairs (A, B, \u2026 without headers).",
    "LAMBDA requires a calculation and at most 253 parameters.",
    "LET requires 1\u2013126 name/value pairs and a final calculation.",
    "LIMIT",
    "Leave calculated-column input slots blank; their saved formulas supply appended values.",
    "Legacy or macro-enabled Excel formats are not supported. Save as XLSX, CSV or native JSON.",
    "Link address",
    "Link addresses cannot contain control characters or backslashes.",
    "Link addresses cannot contain encoded control characters.",
    "Link display text",
    "Link worksheet ID",
    "Linked entity fields must match the declared schema exactly.",
    "Linked worksheet names must be unique.",
    "Links cannot contain embedded credentials.",
    "Local save failed. The previous complete save is unchanged; keep this workbook open and export your unsaved changes.",
    "Local version generation limit reached.",
    "Local version storage exceeds its configured size limit. Export the workbook or use a larger host store; the previous save is unchanged.",
    "MANUAL",
    "MCP result exceeds 2 MB. Ask the host to approve a smaller selection.",
    "Macro-enabled or executable workbook content is unsupported; renaming a file does not convert its format.",
    "Malformed XML source.",
    "Manual page breaks require at most 100 distinct ascending boundaries within each worksheet axis (before row/column 2 or later).",
    "Margins leave no printable page area.",
    "Match the source file extension to the chosen language.",
    "Matching duplicate keys expand beyond 1,000 result rows or reviewable match pairs. Narrow or reconcile the source keys first; no pairs were truncated.",
    "Materiality threshold must be a nonnegative finite number.",
    "Matrix parameters require a static worksheet range.",
    "Merge at most 500 cells.",
    "Merge checking supports at most 10,000 active formulas per worksheet.",
    "Merge exceeds 1,000 intermediate records.",
    "Merge keys cannot repeat a field.",
    "Merged analysis sources are ambiguous.",
    "Merged cells cannot be what-if inputs.",
    "Merged ranges cannot overlap.",
    "Merged ranges exceed the 100,000-cell interchange budget.",
    "Merged ranges exceed the supported import budget (500 cells per merge and 100,000 cells overall).",
    "Merged source cells are ambiguous; use an unmerged headed population.",
    "Merging cells inside a table is unsupported. Convert the table first.",
    "Merging generated analysis output is unsupported.",
    "Migration exceeds 500 touched cells across all commands.",
    "Migration source must be UTF-8 text of at most 256 KiB without nulls.",
    "Minimum cannot exceed maximum.",
    "Missing standard XLSX workbook content type or relationship.",
    "Mixed integer and continuous models are outside this selected Solver subset.",
    "Mixed-content XML tables need explicit extraction.",
    "Model report definition or measure metadata is inconsistent.",
    "Model report range exceeds 20,000 cells.",
    "Model report snapshots cannot be calculation sources. Use the original typed source tables; retained measure caches are not trusted inputs.",
    "Model source formulas require Automatic calculation. In Manual mode use an explicitly reviewed values snapshot; pending or historical formula values cannot enter the model.",
    "More than 100 external workbook targets require an explicit source conversion.",
    "More than 15 significant digits; original text is retained.",
    "More than 50 split fields.",
    "More than one decimal separator.",
    "Moving average needs more observations than its window.",
    "Multi-volume, ZIP64 or oversized XLSX archives are not supported.",
    "Multiple colon-separated VBA statements are unsupported.",
    "Named ranges exceed the 100,000-cell XLSX import/export processing budget. Use native JSON for larger definitions.",
    "Named style IDs and names must be unique within this workbook.",
    "Named style appearance does not match the workbook catalog. Reapply the style from a trusted catalog.",
    "Named style binding is stale or inconsistent with this workbook. Reapply its named style.",
    "Named style binding must contain a namespace, style ID and supported base.",
    "Named style definition must be an object.",
    "Named style identifiers must contain 1\u2013100 characters without control characters.",
    "Names must begin with a letter or underscore and cannot be cell addresses.",
    "Native workbook exceeds the 50 MB import limit.",
    "Native workbook exceeds the 50 MB save limit. Reduce the workbook size before saving.",
    "Navigation requires a valid visible cell.",
    "Nested custom calls and custom calls inside LET/LAMBDA require separate worksheet cells.",
    "No active formula calls use this extension.",
    "No known OLE, ActiveX or embedded-object references were detected. Separately stored objects remain unknown.",
    "No matching workbook link cells.",
    "No numeric amount column found. Include a Debit/Credit, Amount, Net, or Balance column.",
    "No numeric cash flows were supplied.",
    "No rows remain in the report population.",
    "No text cells in the selected range.",
    "No unique worksheet name is available.",
    "No valid contributing amounts. Review the sources and exceptions.",
    "Nonfinite values cannot be used in duplicate review.",
    "Nonlinear multiplication is outside the linear Solver subset.",
    "Normal distribution needs a finite mean and positive standard deviation.",
    "Normalized key text exceeds 2,000,000 characters.",
    "Null setValue semantics are not migrated. Use an explicit ClearContents command.",
    "Number is outside supported precision/range; select Text to retain every digit.",
    "Numeric aggregation encountered text or Boolean values. Add a Types step before aggregating.",
    "Numeric comparisons require a numeric filter value.",
    "Numeric conversion exceeds supported precision or range.",
    "Numeric series require a finite numeric start value.",
    "Numeric series require a finite numeric stop value.",
    "OVERFLOW",
    "Only PNG/JPEG images are supported.",
    "Only Solver can apply changing-cell values.",
    "Only XLTX templates use the native template adapter.",
    "Only a LAMBDA formula can be called with arguments.",
    "Only a candidate that meets the residual tolerance can be applied.",
    "Only a successful reviewed analysis can be saved or applied.",
    "Only an existing linked output can be refreshed.",
    "Only complete A1 cell rectangles can move to a new location.",
    "Only https://, http:// and mailto: links are supported.",
    "Only literal text, finite numbers or booleans are supported in VBA assignments.",
    "Only observed XLSX media may have container metadata.",
    "Only one parameterless VBA Sub may be migrated.",
    "Only selected fields can be renamed.",
    "Only the spill anchor is included; runtime spill output dependencies are not expanded.",
    "Open the saved configuration's source worksheet before loading it.",
    "Opening another workbook requires an explicit namespace replacement or a separate namespace.",
    "Option Explicit requires a declared worksheet/range binding.",
    "Ordered filters require a nonblank number or ISO date.",
    "Ordered filters require a numeric criterion.",
    "Original XLSX exceeds the 10 MiB retention limit.",
    "Original source row identities cannot be changed or removed. Keep the source registry and change only current row associations.",
    "Outline group not found.",
    "Outline groups must be nested or separate; partially overlapping groups are not supported.",
    "Outline groups require unique IDs, valid row or column bounds, and a boolean collapsed state.",
    "Outline groups support at most 8 nested groups.",
    "Output column names must be unique.",
    "Output sheet name must be a valid, unpadded name of 1\u201331 characters.",
    "Output worksheet already exists. Choose a new name; previous successful output is preserved.",
    "Output worksheet names need 1\u201331 characters without control/Excel-reserved characters or leading/trailing apostrophes.",
    "Output-name safety checking exceeds 1,000,000 formula characters. Use a smaller reviewed workbook.",
    "Output-name safety inspection exceeds 1,000,000 formula characters. Use a smaller reviewed workbook.",
    "Overlapping or unlisted XLSX archive entries.",
    "Overlapping validation rules cannot be represented by one Excel rule per cell. Use native JSON or resolve overlapping rules before export.",
    "PDF creation was cancelled.",
    "PDF images exceed 50 MB. Reduce the print area or image sizes.",
    "PDF rendering requires a browser with canvas support.",
    "PNG chunk length or type is invalid.",
    "PNG chunk structure is invalid or excessive.",
    "PNG data or end marker is missing.",
    "PNG encoding failed.",
    "PNG encoding timed out.",
    "PNG exceeds 16 MB. Export SVG or narrow the report.",
    "PNG has duplicate dimensions.",
    "PNG has invalid trailing data.",
    "PNG rendering is unavailable in this browser.",
    "PNG requires an intact first IHDR chunk.",
    "Package link identities must be unique.",
    "Page fitting and scale must be whole numbers.",
    "Page settings changed. Create the PDF again from the updated preview.",
    "Page settings changed. Run the print check again.",
    "Page setup contains an out-of-range number.",
    "Paint at most 500 cells.",
    "Partial-span XLSX page breaks are not supported. Use full-row or full-column breaks.",
    "Password is too long.",
    "Passwords must be text of at most 200 characters.",
    "Paste exceeds the Excel worksheet boundary (XFD1048576).",
    "Paste into at most 500 target cells.",
    "Paste requires a rectangular range.",
    "Pattern checks require pattern settings; legacy threshold/sample/date options do not apply.",
    "Pattern checks support at most 500 selected source cells.",
    "Pattern comparison requires a formula of at most 4,096 characters.",
    "Pattern configuration must be at most 2,000,000 bytes.",
    "Pattern settings require kind patterns.",
    "Pattern source is outside current worksheet dimensions.",
    "Payment timing must be 0 (end) or 1 (beginning).",
    "Period must contain at most 100 characters.",
    "Periods must be greater than zero.",
    "Picture dimensions are invalid or exceed 40 megapixels / 10,000 pixels per side.",
    "Picture evidence digest or byte length does not match the embedded image.",
    "Picture hashing requires at most 2 MB of bytes.",
    "Pictures and shapes do not use population filters.",
    "Pictures in one worksheet may use at most 10 MB.",
    "Pictures must be at most 2 MB.",
    "Pictures must be local PNG/JPEG data of at most 2 MB.",
    "Pictures require canonical local PNG/JPEG base64 data.",
    "Pivot aggregation exceeds the 2,000,000 contribution budget. Reduce value fields, hierarchy depth or source rows.",
    "Pivot definition not found.",
    "Pivot distinct counts exceed 250,000 retained group entries. Reduce groups, value fields or source rows.",
    "Pivot exceeds 50 output columns.",
    "Pivot field sorting requires ascending or descending order.",
    "Pivot keys, attribute and value fields must be distinct.",
    "Pivot population evidence does not reconcile.",
    "Pivot refresh evidence exceeds the 16-million-character fingerprint budget. Use a smaller source.",
    "Pivot view changes can only set filters, sorting, value display and chart settings.",
    "PivotChart requires a supported chart and an existing value field.",
    "PivotCharts are limited to 200 leaf groups. Filter or regroup before charting.",
    "Poisson points must be integers from zero through 1,000; provide a positive rate.",
    "Population membership changed. Preview again before saving or applying.",
    "Positional matching does not use keys.",
    "Preserve the original source before downloading it from the workbook.",
    "Preserved metadata exceeds the XLSX reopen budget.",
    "Preview a merge or unmerge of at most 500 cells.",
    "Preview again before applying.",
    "Preview again.",
    "Preview the current inputs and permissions first.",
    "Preview the current population before saving or applying a view.",
    "Preview the current recipe before publishing.",
    "Print area exceeds the worksheet.",
    "Print assets did not finish loading within 10 seconds. Check the pictures and fonts, then retry.",
    "Print preparation was cancelled.",
    "Print preview supports at most 5,000 cells per print area. Choose a smaller print area.",
    "Print titles must be a single row or column interval.",
    "Provide 2\u201310 examples from distinct source rows.",
    "Provide a bounded source range and scalar field selection.",
    "Provide a complete, supported aging configuration.",
    "Provide a complete, supported materiality screen configuration.",
    "Provide a list of selected ranges to watch.",
    "Provide a valid source identity, name and origin kind.",
    "Provide examples from distinct source rows.",
    "Provide one example value per row, on separate lines.",
    "Provide one reviewed header for every table column.",
    "Qualified name is outside the permitted context.",
    "Query definition not found.",
    "Quote inside an unquoted field; quote the entire field.",
    "Quoted field is not closed.",
    "R1C1 syntax is not resolved.",
    "ROUND digits must be an integer from -12 to 12.",
    "Range exceeds worksheet bounds.",
    "Range is too large. Select at most 100,000 cells.",
    "Range moves cannot repair formula-based conditional rules yet. Remove those rules before moving cells.",
    "Range moves inspect at most 100,000 populated cells.",
    "Range requires a worksheet binding.",
    "Read one to eight explicit source rectangles.",
    "Read-only must be boolean.",
    "Read-only policy must be boolean.",
    "Read-only state must be boolean.",
    "Reading the attachment timed out. Try a smaller file.",
    "Recipe execution evidence exceeds 8,000,000 characters. Retain a native backup and remove obsolete output sheets.",
    "Recipe execution metadata is inconsistent.",
    "Recipe fingerprint exceeds the 16-million-character evidence budget.",
    "Recipe output exceeds workbook cell capacity.",
    "Recipe output must contain finite scalar values.",
    "Recipe preview or authorization changed. Preview again.",
    "Recipe source notes exceed the cell note budget.",
    "Recipe text exceeds the cell limit.",
    "Recipe text or source-lineage budget exceeded. Narrow the source; no evidence was truncated.",
    "Record newline does not match the selected dialect.",
    "Recursive reusable formula: static references are shown once; execution depth is bounded by calculation.",
    "Reference has invalid or out-of-bounds coordinates.",
    "Reference repair exceeds 10 million formula characters. Simplify the workbook before moving cells.",
    "Referenced worksheet is missing.",
    "Refresh uses the saved definition. Publish a new output to change it.",
    "Regression needs more observations than fitted coefficients.",
    "Regression predictors are rank deficient or numerically dependent.",
    "Regression requires Y in the first column and one or more X columns.",
    "Regression requires a nonconstant response.",
    "Relationship keys must be physical source columns, not DAX calculated columns.",
    "Relationships form a cycle. Use a single-direction acyclic model.",
    "Rename each input field only once per step.",
    "Rename fields must be a name map.",
    "Rename separately from changing the definition or comment. Existing scope cannot be changed.",
    "Renaming this table field affects more than 500 cells. Reduce the dependent formulas before renaming it.",
    "Reopen the discussion before adding a reply.",
    "Reordering worksheets with three-dimensional references requires explicit references first.",
    "Repair broken #REF! references before saving a calculated column rule.",
    "Repeated print content exceeds 100,000 rendered cells. Reduce titles, pages or the print area.",
    "Repeated titles leave no space for worksheet data. Reduce the titles, margins or scale.",
    "Replace may change at most 500 cells, including dependent formula changes. Narrow the scope or replace one result.",
    "Replace searches Raw contents or Formulas only. Calculated and formatted results are read-only search modes.",
    "Replacement exceeds the cell text limit.",
    "Replacement text must contain at most 32,767 characters.",
    "Reply",
    "Report calculation is outside the captured range.",
    "Report capture requires automatic, non-iterative calculation. Recalculate and retry.",
    "Report image exceeds 4,096 \xD7 16,384 / 16 megapixels. Narrow the range or export JSON/HTML instead.",
    "Report image exceeds its pixel budget. Narrow the range or export JSON/HTML.",
    "Report image rendering timed out.",
    "Report metadata is not canonical.",
    "Report output not found.",
    "Report ranges cannot be reversed.",
    "Report snapshot integrity check failed.",
    "Reports support at most 500 rows, 20 columns and 2,000 cells.",
    "Requested range is outside the host-approved selection.",
    "Requested worksheet is outside the host-approved selection.",
    "Require issue before first interest and settlement, positive rate and positive par.",
    "Require issue before settlement, positive coupon rate and positive par.",
    "Require settlement before maturity, nonnegative coupon/yield, and positive redemption/price.",
    "Reset choices must be boolean.",
    "Resize a table from its existing top-left header. Moving table source cells is a separate operation.",
    "Resolve all consolidation errors before publishing. No partial result is permitted.",
    "Resolve chart errors before exporting the visual.",
    "Resolve every cleanup error before creating output. No cells were changed.",
    "Resolve model diagnostics and preview again. No partial model result is published.",
    "Resolve the listed exceptions or explicitly choose to retain them unchanged.",
    "Resolved print identifiers exceed Excel's 255-character header/footer limit. Shorten the header, footer or workbook properties before exporting.",
    "Results and exception appendix exceed 5,000 output cells.",
    "Retained XLSX source byte length or encoding changed.",
    "Retained XLSX source checksum does not match its bytes.",
    "Retained report values differ from current source values. Review a fresh preview.",
    "Retained source byte budget exceeded.",
    "Retained source byte length changed.",
    "Retained source checksum mismatch.",
    "Retained structured evidence does not match its original.",
    "Reusable formula evaluation exceeds 1,024 calls.",
    "Reusable formula intermediate values exceed 100,000 cells.",
    "Review and apply a current population filter before visible-only analysis.",
    "Review and approve the complete transformation preview first.",
    "Review and confirm the ambiguous table selection.",
    "Review discussions exceed the per-thread or workbook message limit.",
    "Review every low-confidence cell before importing.",
    "Review message ID",
    "Review message IDs must be unique within the workbook.",
    "Review message times must follow their discussion in order.",
    "Review table headers: resolve blank, duplicate, non-text or formula headers explicitly. Original header values will be retained in native source evidence.",
    "Review thread ID",
    "Review thread IDs must be unique.",
    "Review update time cannot precede creation.",
    "Review worksheet ID",
    "Review worksheet name",
    "Reviewed table headers must be unique, nonblank text.",
    "Reviewer name",
    "Revision",
    "Round-amount increment/tolerance must fit the declared precision; tolerance must be less than half the increment.",
    "Rounded date/time exceeds year 9999.",
    "Row combinations require distinct existing row rules. Sequence gaps are separate interval findings, not missing source rows.",
    "Row count",
    "Row expression nesting exceeds 32 levels.",
    "Row expressions are limited to 200 nodes.",
    "Row heights must be a plain index map.",
    "Row index is outside the worksheet.",
    "Row provenance exceeds its 100,000 identity budget.",
    "Row provenance has reached its 100,000 source-row budget.",
    "Row provenance requires a registry of at most 100,000 source rows.",
    "Rows must have distinct, ascending row numbers.",
    "Rules cannot be changed in this view.",
    "Run a current aging schedule before downloading.",
    "Run a current report first.",
    "Run a current screen before downloading a report.",
    "Runtime reference selection is not resolved; static argument references are only potential dependencies.",
    "SCHEMA",
    "SNAPSHOT_SOURCE",
    "SOURCE",
    "SUMPRODUCT ranges must have matching dimensions.",
    "Safe workbook recovery could not acquire its document lock.",
    "Safe-export prefix would exceed the 32,767-character field limit. Shorten the text before exporting.",
    "Sample seed must contain at most 200 characters.",
    "Sample size must be a whole number from 1 to 500.",
    "Save a recipe version before publishing its output.",
    "Save or export your current unsaved changes before restoring a checkpoint.",
    "Save this workbook before restoring a checkpoint.",
    "Saved assumptions must not repeat a source cell.",
    "Saved baseline inputs and dependency assumptions do not reconcile.",
    "Saved checkpoint metadata is invalid; original storage was preserved.",
    "Saved filter view names and IDs must be unique.",
    "Saved filter views are limited to 20 views and 512,000 metadata characters per workbook.",
    "Saved original field transformation evidence cannot be changed.",
    "Saved original text cleanup evidence cannot be changed.",
    "Saved recipe metadata exceeds 2,000,000 characters.",
    "Saved recipe names and IDs must be unique.",
    "Saved recipe no longer exists.",
    "Saved recipe versions or definition hashes are inconsistent.",
    "Saved reports exceed 16 MB.",
    "Saved selection exceeds worksheet dimensions.",
    "Saved selection must belong to the active visible worksheet.",
    "Saved table source evidence cannot be changed.",
    "Saved what-if IDs must be unique.",
    "Saved what-if baseline is missing.",
    "Saved what-if results must match their named assumptions.",
    "Scalar extension arguments cannot contain arrays.",
    "Scalar formula migration requires one cell; use explicit setFormulas matrices for multiple cells.",
    "Scenario names must be unique, ignoring case.",
    "Search pages require a nonnegative page and 1\u2013200 results per page.",
    "Search switches must be boolean.",
    "Security exceeds 1000 coupon periods.",
    "Select a headed rectangle with 1\u2013500 data rows, 1\u20136 columns and at most 2,000 cells.",
    "Select a matching result before replacing one cell.",
    "Select a populated record or an explicit range to sort.",
    "Select a source table.",
    "Select an empty cell for AutoSum.",
    "Select an existing drawing for an object report.",
    "Select an existing table column.",
    "Select an existing table.",
    "Select at least one data row in addition to the heading row.",
    "Select at least one measure to publish.",
    "Select at least two cells to merge.",
    "Select at least two columns for a matrix.",
    "Select at most 20 chart series.",
    "Select at most 500 cells inside this worksheet for a chart.",
    "Select at most 500 cells to apply a theme.",
    "Select at most 500 cells to remove duplicates.",
    "Select at most 500 source cells and 50 fields.",
    "Select at most 500 source cells for duplicate review.",
    "Select data rows above the table totals row.",
    "Select one cell address.",
    "Select one column for vertical series or one row for horizontal series.",
    "Select one empty total cell.",
    "Select one rectangular range before asking the assistant.",
    "Select one source column containing at most 500 cells.",
    "Select ranges on one worksheet at a time.",
    "Selected raw evidence contains a nonfinite number and cannot be preserved in JSON. Correct it before screening.",
    "Selected source text exceeds the 200,000-character processing budget.",
    "Selection exceeds worksheet dimensions.",
    "Self relationships are unsupported; use distinct role-playing tables.",
    "Sensitivity tables are reports. Only named scenarios can be applied.",
    "Separately qualified range endpoints are not supported for a safe move.",
    "Sequence bounds must be chronological integers on the declared step.",
    "Sequence requires a nonnegative integer or matching prefix/digit/suffix text, and nonblank grouping keys.",
    "Sequence value is outside the declared start/end bounds.",
    "Sequence values or bounds do not align with the declared step.",
    "Series handle fill requires exactly two seed cells in the extension's row or column.",
    "Series handle seeds must evaluate to finite numbers.",
    "Series step must be a nonzero finite number; date steps must be whole numbers.",
    "Set cell must contain an actual calculation formula.",
    "Sheet ID",
    "Sheet IDs and names must be nonempty and unique; sheet names cannot contain Excel-reserved characters.",
    "Sheet cells must be an address map.",
    "Sheet hidden flag must be boolean.",
    "Sheet name",
    "Sheet names must contain 1\u201331 characters and cannot contain \\ / [ ] * ? : .",
    "Sheet names must contain 1\u201331 characters and cannot contain control characters, Excel-reserved characters, or begin or end with an apostrophe.",
    "Slicer selections must be unique.",
    "Solver arithmetic exceeds the supported finite magnitude.",
    "Solver dependencies must be populated numeric cells; blanks and spill children are excluded.",
    "Solver dependencies must contain numbers or supported linear formulas.",
    "Solver dependency budget exceeded.",
    "Solver division requires a nonzero constant denominator.",
    "Solver expression budget exceeded.",
    "Solver expression nesting limit exceeded.",
    "Solver formulas contain a cycle.",
    "Solver range exceeds 500 cells.",
    "Solver syntax budget exceeded.",
    "Sort supports at most 10,000 rows and 100,000 record cells per reviewed operation.",
    "Sorting cannot safely relocate array/spill formulas or existing reference-error tokens. Resolve or separate those formulas first.",
    "Sorting cannot safely rewrite dynamic INDIRECT/OFFSET references. Replace them with explicit references first.",
    "Source and destination date systems differ. Convert and review the source date system before linking.",
    "Source approvals are controlled by your host.",
    "Source body is outside the supported text budget.",
    "Source byte budget exceeded.",
    "Source calculation is unavailable, unauthorized, cyclic or over its resource limit.",
    "Source capture supports at most 50 formula source cells. Use a reviewed values snapshot for a larger calculated population.",
    "Source cell budget exceeded.",
    "Source control totals do not reconcile to the contributing amounts.",
    "Source data or the reviewed plan changed. Preview consolidation again.",
    "Source exceeds the 20,000-token migration limit.",
    "Source extension and format disagree.",
    "Source extension does not match PDF content.",
    "Source extension does not match an HTML document.",
    "Source format and extraction method disagree.",
    "Source formula dependencies produced an array; use a reviewed scalar population.",
    "Source formulas must have deterministic scalar dependencies; volatile, dynamic and array formulas are excluded.",
    "Source headers must be nonblank, unique text. Create a table to normalize blank or duplicate headers.",
    "Source headers must be unique, including case variants.",
    "Source headings must be at most 100 characters.",
    "Source import count exceeded.",
    "Source import is unavailable.",
    "Source inspection does not match its original.",
    "Source provenance metadata budget exceeded.",
    "Source row associations must be a plain index map.",
    "Source row identity capacity has been reached.",
    "Source table IDs must be unique.",
    "Source table count exceeded.",
    "Source table is missing. Review the source mapping.",
    "Source text contains characters SVG cannot represent. Clean the text or export JSON/HTML.",
    "Source text exceeds the 200,000-character processing budget.",
    "Source workbook is missing or has not been approved. Attach and approve the exact workbook identity before refreshing.",
    "Source worksheet is missing or outside the approved worksheet scope.",
    "Specify raw rows or a row count, not both.",
    "Spill discovery exceeds 1,000 possible array formulas.",
    "Spill discovery exceeds 100,000 calculation reads.",
    "Spill reference is inspected at its anchor; runtime extent is not calculated.",
    "Spill result does not match the selected type; the spill is preserved.",
    "Split anchors must be valid zero-based worksheet coordinates.",
    "Split dividers must be between 20% and 80%.",
    "Split one field at a time.",
    "Split output exceeds the worksheet boundary.",
    "Split output may contain at most 500 cells. Select fewer source rows or fields.",
    "Split panes and frozen panes cannot be enabled together.",
    "Split panes require divider percentages and valid row/column anchors.",
    "Stale calculation delta",
    "Start a new chat to continue. This chat has reached its message limit.",
    "Start the custom validation formula with =.",
    "Static comparison supports 20,000 permitted stored cells across both workbooks.",
    "Static formula inspection exceeds the 2,000,000-character work budget.",
    "Stop errors cannot be overridden. Correct the invalid values.",
    "Stop if true is not supported for color scales, data bars, or icon sets.",
    "Stop if true must be boolean.",
    "Stored #REF! token identifies a broken reference.",
    "Stored workbook snapshot does not satisfy workbook validation.",
    "Stored workbook snapshot failed its integrity check.",
    "Stored workbook snapshot is incomplete.",
    "Strict or malformed conditional icon thresholds are not supported. Use inclusive thresholds or native JSON.",
    "Structural edits cannot safely update 3-D worksheet references.",
    "Structural edits cannot safely update 3-D worksheet references. Replace the 3-D reference with explicit worksheet references first.",
    "Structural edits cannot safely update external workbook references.",
    "Structural preview can verify up to 5,000 formula cells. Reduce or separate the formula population before changing worksheet structure.",
    "Structured reference could not be inspected.",
    "Structured table reference could not be resolved; table/column or syntax may be unsupported.",
    "Style colors must use hexadecimal colors.",
    "Subtotals must be boolean.",
    "Suggest patterns and explicitly choose one before adding the inferred step.",
    "TYPE",
    "Tab colors must use a six-digit hexadecimal color.",
    "Table IDs and names must be unique.",
    "Table append/resize may change at most 500 cells.",
    "Table append/resize would overwrite adjacent cells or their formatting. Move that content first.",
    "Table column IDs and names must be unique.",
    "Table columns or totals do not fit the table range.",
    "Table headers must be nonblank text. The header edit was not applied.",
    "Table headers must be unique, nonblank text. The header edit was not applied.",
    "Table headers, totals and calculated columns cannot be changing cells.",
    "Table headers, totals and calculated columns cannot be what-if inputs.",
    "Table names and IDs must be unique across the workbook.",
    "Table not found.",
    "Table settings or the current selection changed. Preview again.",
    "Table, query and pivot ranges are limited to 100,000 cells including blank gaps. Select the used data range.",
    "Tables cannot overlap.",
    "Tables must be a bounded list.",
    "Target and bounds must be finite numbers, with lower bound smaller than upper bound.",
    "Text attachments must use UTF-8 encoding. Save the file as UTF-8 and try again.",
    "Text automation source is limited to 256 KiB.",
    "Text criteria require a text value.",
    "Text filters require a text criterion.",
    "Text matching requires a Text condition.",
    "Text replace needs nonempty text to find and a text replacement.",
    "Text to Columns permissions must be boolean.",
    "That output name is already referenced by workbook formulas or formula metadata. Choose a different name to avoid reconnecting an unresolved source reference.",
    "The Excel worksheet has an unsupported pageOrder value. Import stopped without changing the current workbook.",
    "The Goal Seek preview or host context changed. Solve and review again.",
    "The Name Box accepts cell addresses and static range names. Computed, external, or three-dimensional destinations cannot be selected.",
    "The PDF exceeds 50 MB. Reduce the print area or image sizes.",
    "The PDF page plan changed. Reopen print preview and try again.",
    "The PDF rendering frame could not load.",
    "The PDF rendering frame could not load. Check the embedding host's content policy.",
    "The PivotChart contains a calculation error. Resolve zero denominators or invalid source calculations.",
    "The Set cell formula does not depend on the selected changing cell.",
    "The XLSX archive contains incomplete entries.",
    "The XLSX archive exceeds the safe expansion limit (50 MB or 200:1 ratio).",
    "The XLSX archive expands beyond 50 MB.",
    "The XLSX archive expansion does not match its directory.",
    "The XLSX archive is incomplete or has trailing data.",
    "The XLSX workbook contains corrupt or unsupported records. Open the original in Excel, repair it if needed, and save a clean XLSX copy.",
    "The address or defined range name was not found in the current worksheet scope.",
    "The aggregate missing-value count exceeds the exact integer reporting limit. Run fewer sequence rules.",
    "The aging configuration/report is not valid JSON.",
    "The aging configuration/report must be at most 2,000,000 bytes.",
    "The aging control total exceeds finite numeric capacity. Reduce the selected population or magnitude.",
    "The aging report exceeds 2,000,000 bytes. Select a smaller population.",
    "The allowed worksheet context is invalid.",
    "The analysis output exceeds 100,000 cells.",
    "The analysis range must match the declared filter population.",
    "The appended table exceeds the worksheet boundary.",
    "The approved automation scope changed. Review again.",
    "The approved worksheet does not exist.",
    "The approved worksheet is protected.",
    "The attachment could not be read. Select it again.",
    "The authorized worksheet context is invalid.",
    "The browser changed the PDF page resolution. Use browser printing instead.",
    "The browser could not render a PDF page.",
    "The browser could not render this report image. Export SVG or HTML.",
    "The calculation worker could not run. Check the host's worker policy, then retry; inputs are intact.",
    "The calculation worker returned an invalid result. Inputs are intact; retry the calculation.",
    "The calculation worker returned no usable result. Inputs are intact; retry the calculation.",
    "The cell address exceeds Excel dimensions.",
    "The changing worksheet is protected. Unprotect it before Goal Seek.",
    "The chart no longer exists.",
    "The chart worksheet no longer exists.",
    "The chosen cell does not match this search. Find it again before replacing.",
    "The cleanup preview or authorized context changed. Preview again.",
    "The cleanup recipe exceeds the 4,096-character formula budget.",
    "The cleanup settings changed. Preview again.",
    "The column type preview changed. Preview again.",
    "The comparison worksheet access context is invalid.",
    "The complete evidence report exceeds 2,000,000 bytes. Select a smaller source population.",
    "The complete pattern report exceeds 2,000,000 bytes. Select fewer cells/rules; results are never silently truncated.",
    "The complete source evidence exceeds 2,000,000 bytes. Select a smaller population.",
    "The configured aging buckets do not cover a valid source row.",
    "The control total exceeds finite numeric capacity. Reduce the population or magnitude.",
    "The copied range exceeds the worksheet.",
    "The current approved source changed or editing is unavailable. Request a fresh answer before applying changes.",
    "The current approved source changed. Request a fresh answer.",
    "The data model exceeds 16,000 source cells.",
    "The date step exceeds safe integer limits.",
    "The declared dimension has no valid cell range.",
    "The defined name belongs to a worksheet outside the selected inspection context.",
    "The defined name is outside the authorized inspection context.",
    "The defined name no longer exists.",
    "The defined name was not found. Reopen Name Manager.",
    "The defined range name contains a circular reference.",
    "The defined-name chain exceeds the 32-name navigation limit.",
    "The defined-name expansion limit was reached.",
    "The derived column name already exists.",
    "The derived workbook would exceed the native file budget.",
    "The derived worksheet would exceed the workbook's stored-cell budget.",
    "The destination calculation includes an unresolved or unavailable dependency.",
    "The destination contains a saved watch or discussion anchor. Choose another range to preserve its evidence context.",
    "The destination contains an existing review discussion. Choose a destination without anchored discussions to preserve its evidence context.",
    "The destination contains cells or their formatting. Choose an empty range or new worksheet.",
    "The destination is inside or immediately after the selected block; this would not move anything.",
    "The destination is outside the available worksheet scope.",
    "The destination worksheet is hidden. Make it visible before navigating.",
    "The directional seed must contain a number or date, or supply a start value.",
    "The discussion could not be found. Reopen the review tools.",
    "The drawing anchor is hidden. Explicit authorization is required.",
    "The drawing no longer exists.",
    "The drawing preview is stale. Reopen it or select the object again.",
    "The drawing selection changed. Reopen this dialog.",
    "The duplicate preview or host context changed. Preview and review again.",
    "The duplicate review host context is invalid.",
    "The editor changed before the restore completed. Your current local changes were retained.",
    "The export range must match the declared filter population.",
    "The extension adapter failed. No workbook changes were applied.",
    "The extension formula is outside its approved worksheet scope.",
    "The extension input range exceeds its declared scalar/cell budget.",
    "The extension is missing or disabled.",
    "The extension returned an invalid response envelope.",
    "The field does not match the explicit numeric separators.",
    "The file bytes changed while reading the picture.",
    "The file contents do not match a JPEG picture.",
    "The file contents do not match a PNG picture.",
    "The file has no records. Select a populated delimited file.",
    "The filename extension does not match the workbook content. Restore its original extension or convert it in the source application; renaming is not conversion.",
    "The fill handle may extend into at most 500 cells.",
    "The fill range exceeds the worksheet.",
    "The fill-handle target exceeds the worksheet.",
    "The fill-handle target must contain the complete source range.",
    "The final worksheet cannot be deleted.",
    "The financial equation overflowed or has no finite result. Review rates, signs and periods.",
    "The format inspection changed or is untrusted. Inspect the original again.",
    "The formula contains an error token.",
    "The formula destination changed.",
    "The formula watch host context contains an unknown worksheet.",
    "The formula-cell scan limit was reached; additional dependents may exist.",
    "The function is outside the local supported catalog; its static references may be incomplete.",
    "The generated worksheet has an invalid row number.",
    "The generated worksheet has invalid column metadata.",
    "The generated worksheet is missing its data section.",
    "The handle series would overflow.",
    "The host calculation changed. Read the current calculation snapshot before continuing.",
    "The host changed this workbook to read-only. No structure change was applied.",
    "The host does not allow the isolated PDF rendering document.",
    "The host has not approved this exact manifest and all of its permissions.",
    "The host has not granted all declared extension permissions.",
    "The host scope contains an unknown worksheet.",
    "The identifier is not a defined name or supported reference.",
    "The inferred output would exceed 50 fields.",
    "The inferred pattern no longer matches every reviewed example.",
    "The inferred pattern reads a field outside its approved inference columns.",
    "The inferred result exceeds the cell text limit.",
    "The inferred rule refers to a missing field.",
    "The inferred separator is absent.",
    "The inferred token is absent.",
    "The intermediate result exceeds 5,000 cells.",
    "The legacy saved workbook is invalid; original storage has been preserved.",
    "The link contains invalid encoded characters.",
    "The linked output was removed.",
    "The literal matrix dimensions do not match the target range.",
    "The materiality screen configuration version is unsupported.",
    "The materiality screen host context is invalid.",
    "The materiality screen supports at most 500 selected source cells.",
    "The merge preview changed. Create a fresh preview.",
    "The migration plan is invalid, modified or untrusted. Create and review a fresh plan.",
    "The model query exceeds 1,000 groups.",
    "The model result exceeds 20,000 output cells.",
    "The model supports at most 16 DAX calculated columns.",
    "The model, workbook or host permissions changed. Preview again.",
    "The move exceeds workbook history limits. Export the workbook before starting a new workpaper.",
    "The move intersects a dynamic-array footprint. Convert it to ordinary values or move cells outside the array.",
    "The move intersects a generated PivotTable or query output. Move ordinary cells instead.",
    "The move intersects a validation or conditional-formatting rule. Remove or relocate the rule first.",
    "The move no longer matches its preview. Cut the source again.",
    "The move would change more than 500 metadata references. Narrow the source.",
    "The move would create, remove, or resize a dynamic-array footprint. Convert the affected array to ordinary values before moving cells.",
    "The moved cells would exceed the worksheet boundary.",
    "The name dependency chain exceeds 32 levels.",
    "The name dependency scan exceeds the supported budget. Split the workbook before changing names.",
    "The name preview changed. Review the operation again.",
    "The native workbook is not valid JSON.",
    "The new output would exceed the workbook's stored-cell budget.",
    "The new totals row already has a saved watch or review discussion. Preserve or remove that destination anchor before appending.",
    "The new totals row would overwrite existing cells. Include an empty final row in the resized range.",
    "The normalized link is too long.",
    "The numeric series would overflow.",
    "The objective must be a linear formula cell.",
    "The old totals row was edited. Preserve or move those values before resizing it.",
    "The original saved workspace changed during recovery.",
    "The output destination contains existing cells. Choose an empty destination or a new sheet.",
    "The output destination overlaps a table.",
    "The output destination overlaps an existing merged area or analysis result. Choose another empty destination.",
    "The output does not fit the destination worksheet.",
    "The output exceeds 1,000 data rows or 20,000 cells including headers and provenance.",
    "The output name would reconnect an existing unresolved worksheet reference. Choose a different name.",
    "The output sheet name already exists. Choose a new sheet name.",
    "The output sheet needs a valid Excel worksheet name of 1\u201331 characters.",
    "The output worksheet is protected.",
    "The output worksheet name is already in use.",
    "The output worksheet name is already referenced by workbook formulas or formula metadata. Choose another name to avoid reconnecting an unresolved source reference.",
    "The output worksheet name is already used.",
    "The output worksheet name is already used. Choose another name.",
    "The password does not match.",
    "The pasted cells would exceed the worksheet boundary.",
    "The pattern configuration is not valid JSON.",
    "The picture could not be decoded.",
    "The picture data is invalid.",
    "The pivot grouping would exceed the 100,000-cell output budget. Reduce group fields or filter the source.",
    "The preview or authorized context changed. Preview again.",
    "The preview or host context changed. Preview and review again.",
    "The preview, workbook or host permissions changed. Preview again.",
    "The preview, workbook or permissions changed. Preview the report again.",
    "The print area has no visible cells. Unhide its rows or columns, or choose another print area.",
    "The print content cannot fit the selected paper. Reduce the print area or titles.",
    "The print window was closed. Reopen print preview to retry.",
    "The qualified defined name does not exist.",
    "The qualified reference cannot be safely migrated.",
    "The query output exceeds 100,000 cells.",
    "The quoted field is not closed.",
    "The range exceeds the worksheet boundary.",
    "The range would exceed the Excel worksheet boundary.",
    "The rate guess is outside the supported search interval.",
    "The rate must be greater than -100%.",
    "The recipe differs from the saved version. Save a new version or reload the selected version.",
    "The reference is outside Excel's address bounds or is invalid.",
    "The reference qualifier cannot be safely migrated.",
    "The reference syntax is outside the inspected subset.",
    "The referenced worksheet does not exist.",
    "The report destination changed. Export again.",
    "The report is outside the current authorized source context.",
    "The report range must match its filter range.",
    "The report range must match the drawing source.",
    "The report source is missing or hidden.",
    "The report source is outside the authorized host context.",
    "The report source worksheet is hidden.",
    "The requested extension is unavailable.",
    "The resized table does not fit the worksheet or processing budget.",
    "The result needs 1\u201350 distinct, nonblank column names of at most 128 characters.",
    "The result would exceed the 100,000 populated-cell workbook budget.",
    "The resulting formula has an error or obstructed array. No results were applied.",
    "The reusable definition recurs; references are shown once.",
    "The review time cannot precede the existing discussion.",
    "The saved analysis is missing or outside the authorized context.",
    "The saved analysis is outside the authorized context.",
    "The saved analysis no longer exists.",
    "The saved configuration belongs to another worksheet.",
    "The saved local workbook exceeds this store's size limit. Preserve it for export or increase the configured limit.",
    "The saved version changed again. Refresh the comparison before reloading.",
    "The saved workbook cannot be read. Its original storage has been preserved; export in-memory work before choosing another namespace.",
    "The screen configuration/report is not valid JSON.",
    "The screen configuration/report must be at most 2,000,000 bytes.",
    "The screen excludes volatile, dynamic-reference and array/spill formulas. Use supported scalar values or formulas.",
    "The screen requires scalar fields; a formula produced an array/spill result.",
    "The search selection exceeds the worksheet.",
    "The selected cell is outside the visible, authorized inspection context.",
    "The selected cells are outside the visible, authorized worksheet context.",
    "The selected cells exceed the worksheet boundary.",
    "The selected drawing no longer exists.",
    "The selected formulas exceed the bounded calculation inspection length.",
    "The selected range does not intersect a merged area.",
    "The selected source or destination intersects a generated array. Convert its results to ordinary values explicitly before splitting.",
    "The selected source table no longer exists. Review the recipe source.",
    "The selection has more than 50 possible array anchors. Choose a new output sheet or a smaller reviewed region.",
    "The series date is not a valid calendar date.",
    "The series range exceeds the worksheet.",
    "The shape label does not fit the report image. Enlarge the source shape before publishing.",
    "The shifted table or analysis exceeds the 100,000-cell range budget.",
    "The solution violates a changing-cell validation rule.",
    "The sort exceeds the bounded formula-reference review budget. Reduce or separate the workbook before sorting.",
    "The sort selection is outside the worksheet.",
    "The sort source is outside the host-approved worksheets.",
    "The sort would change more than 100,000 cells including dependent references.",
    "The source drawing no longer exists.",
    "The source inspection changed or is untrusted. Inspect the original again.",
    "The source is outside the authorized host context.",
    "The source is outside the host-approved worksheets.",
    "The source must contain at least one data row after its header.",
    "The source range includes hidden content. Explicitly include hidden content only within authorized host scope.",
    "The source range is outside the visible, authorized host context. Include hidden only within permitted host scope.",
    "The source selection or preview changed. Preview and review again.",
    "The source table no longer exists.",
    "The source worksheet no longer exists.",
    "The split preview or authorized context changed. Preview again.",
    "The split would exceed 50 fields.",
    "The split would exceed the workbook's 100,000 stored-cell budget.",
    "The starting input and both bounds must satisfy the variable's data validation rules.",
    "The starting input must be inside the supplied bounds.",
    "The statistical result exceeds finite numeric precision. Rescale the sample.",
    "The stop value is opposite the series step direction.",
    "The table overlaps another table or exceeds the cell budget.",
    "The table preview was changed. Preview again.",
    "The table, field, or structured-reference form could not be resolved.",
    "The target cell is hidden. Unhide its row/column or expand its outline before navigating.",
    "The target rectangle must be a whole multiple of the copied rows and columns.",
    "The target workbook history is full. Export a saved copy before continuing.",
    "The template content type could not be converted safely.",
    "The total missing-value count exceeds the exact integer reporting limit. Split sequence groups into smaller checks.",
    "The total table footprint exceeds 100,000 cells.",
    "The totals row was edited. Preserve its values before appending.",
    "The totals row would overwrite existing cells.",
    "The transformation preview is stale or altered. Preview again.",
    "The transformed value exceeds the cell text limit.",
    "The value is too short for the inferred character extraction.",
    "The very-hidden flag must be boolean.",
    "The watch preview or host context changed. Review the action again.",
    "The what-if preview or host context changed. Preview and review again.",
    "The workbook changed after Cut. Cut the source again before moving it.",
    "The workbook changed after duplicate review. Preview again before Apply.",
    "The workbook changed after page setup opened. Close and reopen page setup.",
    "The workbook changed after preview. Preview and review the analysis again.",
    "The workbook changed after preview. Preview and review the populations again.",
    "The workbook changed after the Goal Seek preview. Solve and review again.",
    "The workbook changed after the cleanup preview. Preview again.",
    "The workbook changed after the name preview. Review the operation again.",
    "The workbook changed after the split preview. Preview again.",
    "The workbook changed after the table preview. Preview again.",
    "The workbook changed after the watch preview. Review the action again.",
    "The workbook changed after these sizes were opened. Close the dialog and reopen it before applying sizes.",
    "The workbook changed after this move preview. Cut the source again before moving it.",
    "The workbook changed after this preview. Close and reopen Themes & Cell Styles.",
    "The workbook changed after this preview. Create a new preview before applying it.",
    "The workbook changed after this preview. Refresh the preview before applying.",
    "The workbook changed while checking the password. Try again with the current workbook.",
    "The workbook changed while the file was being read. Import again to replace the current revision.",
    "The workbook changed while the file was being read. Open the file again.",
    "The workbook changed while this name was being edited. Choose New name or reopen Edit to review current content.",
    "The workbook changed while this tool was open. Reopen the tool and try again.",
    "The workbook changed. Close this preview and review the current cells again.",
    "The workbook changed. Compare the checkpoint again before restoring.",
    "The workbook changed. Find or preview again before using these results.",
    "The workbook changed. Preview column types again.",
    "The workbook changed. Preview drill-through again.",
    "The workbook changed. Preview pivot filters again.",
    "The workbook changed. Preview the analysis again.",
    "The workbook changed. Preview the import again.",
    "The workbook changed. Preview the model again.",
    "The workbook changed. Preview the report again.",
    "The workbook changed. Preview the source again.",
    "The workbook changed. Preview the transformation again.",
    "The workbook changed. Reopen the filter tool before saving a view.",
    "The workbook changed. Review a fresh migration plan.",
    "The workbook changed. Review the conversion again.",
    "The workbook changed. Review the export again.",
    "The workbook changed. Review the file again.",
    "The workbook changed. Review the fill again.",
    "The workbook changed. Save the checkpoint again for the current revision.",
    "The workbook content does not match the template conversion route.",
    "The workbook is read-only.",
    "The workbook or calculation run changed. Review current calculation results before exporting evidence.",
    "The workbook or host permissions changed after the sort preview. Preview again.",
    "The workbook or permissions changed. Preview links again.",
    "The workbook or permissions changed. Preview the source again.",
    "The workbook or permissions changed. Review extensions again.",
    "The workbook or permissions changed. Review the conversion again.",
    "The workbook or review changed. Preview the import again.",
    "The workbook or selected context changed. Request a new external proposal.",
    "The workbook or selection changed. Close and reopen Format Cells.",
    "The workbook or selection changed. Close and reopen Themes & Cell Styles.",
    "The workbook or selection changed. Close and reopen the review tools.",
    "The workbook or selection changed. Close and reopen this preview.",
    "The workbook or selection changed. Preview again.",
    "The workbook or source range changed. Review headers again.",
    "The workbook reference scan limit was reached.",
    "The workbook revision limit has been reached.",
    "The workbook structure password does not match.",
    "The worker result arrived after the calculation deadline and was discarded.",
    "The worker result contains stale or duplicate formula inputs.",
    "The worker result does not match the requested workbook.",
    "The worksheet changed or was deleted. Close and reopen page setup.",
    "The worksheet no longer exists.",
    "The worksheet or selection changed. Close and reopen page setup.",
    "The worksheet selected for watching no longer exists.",
    "The worksheet selection changed after the watch preview. Review the selection again.",
    "The worksheet was not found.",
    "The zoom changed. Start the drawing move again.",
    "Theme colors must use six-digit hexadecimal colors.",
    "There is no room for a totals row.",
    "This PivotTable is stale or unverified. Refresh it before drilling or charting; historical totals cannot be matched to current source rows.",
    "This action changes more than 500 cells. Use smaller selections or fewer bound cells.",
    "This analysis requires exactly one numeric column.",
    "This archive is not a supported Excel workbook.",
    "This artifact supports inventory only. Export/import UTF-8 VBA or Office Scripts source for migration.",
    "This browser cannot coordinate safe workbook saves between tabs.",
    "This calculator cannot distinguish error values from literal text.",
    "This calculator cannot evaluate conditional formulas.",
    "This checkpoint is unavailable or has expired from the bounded local history.",
    "This command requires one rectangular range. Click a cell or enter a single range to continue.",
    "This conditional formula uses unqualified table fields across a table boundary. Split or remove its applied range before renaming the field.",
    "This discussion is retained for review because its target was deleted.",
    "This edit moves a manual print break beyond the worksheet. Remove the break first.",
    "This edit removes every repeated print title. Clear or change print titles first.",
    "This edit would make outline groups overlap or duplicate. Ungroup the affected rows or columns first.",
    "This edit would remove a chart header row. Adjust or delete the chart first.",
    "This edit would remove a table totals row. Turn off table totals first.",
    "This edit would remove all chart data. Adjust or delete the chart first.",
    "This edit would remove chart category labels. Adjust or delete the chart first.",
    "This edit would remove the entire chart source. Adjust or delete the chart first.",
    "This edit would remove the entire print area. Change the print area first.",
    "This embedding host has not connected multiple-area selection.",
    "This expression requires a finite number. Add an explicit Types step.",
    "This expression requires text. Add an explicit Types step.",
    "This file is not a supported XLSX archive.",
    "This filter requires an explicit typed value.",
    "This formula was outside the formula scan budget.",
    "This host is read-only. Calculation settings cannot be changed.",
    "This is the current workbook identity. Use the current workbook source instead.",
    "This layout exceeds 100 pages. Reduce the print area or fit it to fewer pages.",
    "This link belongs to a different workbook. Publish a new output after reviewing the source.",
    "This list validation cannot be represented by an inline XLSX list. Retain the native workbook or use shorter values without commas, quotes or line breaks.",
    "This local namespace already contains a workbook. Reload it before saving; unsaved work remains in memory.",
    "This move would displace a table header or totals row. Move the complete table or only its data rows.",
    "This move would displace an analysis source header. Move the whole source or only its data rows.",
    "This move would displace chart headers or category labels. Move the complete chart source or only its data rows or value columns.",
    "This named style no longer exists.",
    "This operation exceeds the 100,000 cell-change history limit.",
    "This provider does not support conversation history. Start a new chat or use a chat-enabled host adapter.",
    "This provider does not support file attachments. Remove the files or connect a chat-enabled provider.",
    "This provider does not support file attachments. Remove the files or use a chat-enabled host adapter.",
    "This saved filter view no longer exists.",
    "This saved view cannot be changed here.",
    "This search exceeds 100,000 populated cells. Narrow its scope.",
    "This search exceeds the text-comparison budget. Narrow the selection or simplify the search.",
    "This selection contains more than 256 separate areas. Narrow the selection.",
    "This sheet is protected. Unprotect it before changing its contents.",
    "This sort preview changed or was restored from another session. Preview again.",
    "This source exceeds the 2 MB artifact-retention limit. Keep the original file separately.",
    "This source inventory has already been retained.",
    "This structural preview was changed or restored from another session. Create a fresh preview before applying.",
    "This style belongs to another workbook. Import or remap its style catalog before pasting the binding.",
    "This theme preview changed. Create a fresh preview before applying it.",
    "This value cannot be converted without guessing.",
    "This watch cannot be navigated in the current visible, authorized context.",
    "This workbook has no retained XLSX source. Open an XLSX file to retain its original bytes.",
    "This workbook has no retained alternate-format original.",
    "This workbook is read-only or protected.",
    "This workbook is read-only.",
    "This workbook is read-only. A derived worksheet cannot be created.",
    "This workbook is read-only. Goal Seek cannot propose a writable input.",
    "This workbook is read-only. Sorting cannot change its records.",
    "This workbook is read-only. Watch inspection remains available.",
    "This worksheet has dependencies. Review them and explicitly confirm deletion.",
    "This worksheet has no consolidation record.",
    "This worksheet is read-only or protected.",
    "This worksheet is read-only.",
    "This worksheet is read-only. Unlock it before changing outlines.",
    "Three-dimensional references cannot be safely migrated.",
    "Three-dimensional references require separate review and are not expanded.",
    "Three-dimensional worksheet references are not expanded.",
    "Three-dimensional worksheet references must be replaced with explicit local references before moving cells.",
    "Threshold must be a finite, nonnegative number.",
    "Tickmark",
    "Time bounds must be fractions from 0 up to, but excluding, 1.",
    "Timeline date exceeds year 9999.",
    "Timeline requires ordered YYYY-MM-DD bounds, a period and an explicit date system.",
    "Tolerance must be a finite positive absolute error.",
    "Too many matching cells. Narrow the search.",
    "Trailing JSON content.",
    "Transformed output type does not match its value.",
    "Transpose and skip blanks must be boolean.",
    "Typed XML integer is out of range.",
    "Unable to inspect corrupt alternate workbook archive.",
    "Unable to inspect corrupt automation archive.",
    "Unable to preserve XLSX package metadata safely.",
    "Unable to read the generated template copy.",
    "Unbalanced Solver parentheses.",
    "Unbalanced or malformed formula syntax.",
    "Unbalanced table references cannot be safely migrated.",
    "Unclosed VBA string.",
    "Unclosed block comment.",
    "Unclosed string literal.",
    "Undo cannot change workbook structure protection. Use Manage worksheets to unlock it.",
    "Unexpected End Sub.",
    "Unexpected text after a closing quote.",
    "Unexpected text after a closing quote. A closing quote must be followed by the delimiter or end of the cell.",
    "Unhide the target worksheet with Manage worksheets before navigating to it.",
    "Unhide this sheet before selecting its cells.",
    "Unhide this worksheet before sorting its records.",
    "Unknown automation artifact kind.",
    "Unknown automation migration adapter.",
    "Unknown calculated-table operation field.",
    "Unknown calculated-table operation.",
    "Unknown citation.",
    "Unknown clear mode.",
    "Unknown defined-name field.",
    "Unknown diagnostic adapter or operation.",
    "Unknown extension operation.",
    "Unknown formula watch action.",
    "Unknown link operation.",
    "Unknown name operation.",
    "Unknown or already applied extension preview.",
    "Unknown or already applied extension results.",
    "Unknown or already applied workbook link preview.",
    "Unknown or already executed extension input review.",
    "Unknown or repeated original field source locator.",
    "Unknown recipe request field.",
    "Unknown row expression operation.",
    "Unknown source review cell.",
    "Unlock the workbook structure and protected worksheets before restoring a version.",
    "Unmerge cells in the filter range first; records cannot be split by filtering.",
    "Unmerge covered target cells before filling them.",
    "Unmerge fewer areas; this preview supports at most 500 affected cells.",
    "Unmerge overlapping areas first.",
    "Unmerge selected cells before changing border edges.",
    "Unmerge selected cells before removing duplicates.",
    "Unmerge the approved source before pattern checks.",
    "Unmerge the intersecting cells before applying border edges or presets.",
    "Unmerge the intersecting cells before changing named-style border edges.",
    "Unmerge the selected source before duplicate review so row boundaries are unambiguous.",
    "Unmerge the selected source before reshaping fields.",
    "Unmerge the source and destination before moving cells.",
    "Unmerge the source before cleaning text.",
    "Unmerge the source before splitting text.",
    "Unmerge the source range before screening so every record has unambiguous fields.",
    "Unmerge the target area before pasting a block of cells or editing a covered cell.",
    "Unpaired Unicode surrogate cannot be exported; correct the text first.",
    "Unpivot would exceed 1,000 intermediate rows; narrow the selected fields or source.",
    "Unprotect the workbook and source worksheet before changing saved filter views.",
    "Unprotect the worksheet before changing conditional formats.",
    "Unprotect this sheet before changing its protection settings.",
    "Unrecognized expression for an aggregate measure; only DAX measures may contain it.",
    "Unsafe JSON number; encode large identifiers as strings.",
    "Unsafe or duplicate JSON field.",
    "Unsafe or duplicate XLSX archive path.",
    "Unsupported JPEG pixel encoding.",
    "Unsupported PNG pixel encoding.",
    "Unsupported Solver expression.",
    "Unsupported Solver syntax. Use numeric references, arithmetic, SUM, AVERAGE or SUMPRODUCT.",
    "Unsupported VBA object or dynamic reference. Use a literal worksheet/range in the current workbook.",
    "Unsupported VBA statement/API. Loops, expressions, calls, events, files, COM and external operations are not migrated.",
    "Unsupported XLSX compression or ZIP64 metadata.",
    "Unsupported XLSX page order. Import stopped to preserve the source layout.",
    "Unsupported XLSX path encoding.",
    "Unsupported alternate workbook content type.",
    "Unsupported alternate workbook content-type namespace.",
    "Unsupported alternate workbook package identity.",
    "Unsupported audit analysis.",
    "Unsupported automation language.",
    "Unsupported cell border.",
    "Unsupported cell format.",
    "Unsupported chart type.",
    "Unsupported date policy. Use the Gregorian calendar, 1900 or 1904 serial system, MDY or DMY text dates, and a valid IANA time zone (or UTC).",
    "Unsupported font family.",
    "Unsupported or corrupt local workbook envelope; stored data was preserved.",
    "Unsupported or duplicate workbook relationship.",
    "Unsupported or mismatched OpenDocument workbook identity.",
    "Unsupported or missing worksheet records. Save a standard XLSX copy with ordinary worksheets.",
    "Unsupported or trailing Solver formula syntax.",
    "Unsupported print page order. Choose Over, then down or Down, then over.",
    "Unsupported shape.",
    "Unsupported string escape.",
    "Unsupported workbook content type. Templates and alternate formats must be saved as XLSX.",
    "Unsupported workbook schema. Expected schemaVersion 1.",
    "Unsupported, encrypted or corrupt legacy workbook. Open the original in its source application and save an unencrypted XLSX copy.",
    "Unterminated formula text cannot be inspected completely.",
    "Unterminated formula text; link inspection is incomplete.",
    "Unused external sources are not accepted as grounding.",
    "Use 1\u201320 checkpoints and a local envelope size between 1 KiB and 50 MiB.",
    "Use Rename to change a name. Existing name scope cannot be changed; create a separate name instead.",
    "Use SUM, AVERAGE or SUMPRODUCT to reduce a range.",
    "Use UTF-8 XML.",
    "Use Values paste for formulas copied from a different workbook.",
    "Use a canonical cell address.",
    "Use a nonempty filter view name of up to 80 characters.",
    "Use a number with the selected decimal/group separators; currency, percentages and scientific text require a separate cleanup.",
    "Use a recipe name of 1\u2013128 characters.",
    "Use a row expression of 1\u20131,000 characters.",
    "Use a table name beginning with a letter or underscore, without spaces or a cell-address name.",
    "Use an escaped newline inside a string literal.",
    "Use at most 100 calculated fields.",
    "Use at most 100 date groups.",
    "Use at most 100 derived fields.",
    "Use at most 100 source filters.",
    "Use mailto: with one email address and no additional headers.",
    "Use the ordinary transaction path for the same workbook.",
    "VBA binding type does not match its declaration.",
    "Validation approval is stale or invalid. Review the complete batch again.",
    "Validation inspection is stale or unknown. Inspect existing values again.",
    "Validation rule IDs must be nonempty and unique.",
    "Value filters allow up to 1,000 finite typed values.",
    "Variable bounds must be ordered; binary bounds must lie within zero and one.",
    "Version comparison requires the same workbook identity.",
    "Web links require a hostname.",
    "Weekend weekdays must be distinct (0=Sunday through 6=Saturday).",
    "What-if analysis excludes volatile, dynamic-reference, array/spill and unsupported reusable-function dependencies.",
    "What-if analysis requires a complete dependency closure within 500 cells, 5,000 scanned formulas and depth 8.",
    "What-if analysis requires populated static dependencies; blank or possible spill-child cells are unsupported.",
    "What-if analysis supports at most 500 dependency cells across all results.",
    "What-if cell references must not repeat.",
    "What-if evaluation would exceed 50,000 dependency-cell work units. Reduce the values, scenarios or dependencies.",
    "What-if host scope contains an unknown worksheet.",
    "What-if work counts do not reconcile.",
    "Windows-1252 does not support a BOM; disable BOM or choose UTF-8.",
    "Workbook ID",
    "Workbook ID and name are required.",
    "Workbook changed after preview. Preview the recipe again.",
    "Workbook changed while copying. Copy again.",
    "Workbook changed while pasting. Try again.",
    "Workbook changed. Review the batch again.",
    "Workbook content changed without a new revision. Commit the edit before saving.",
    "Workbook context is stale. Read current context before proposing changes.",
    "Workbook data or approved selection changed. Request a new context and proposal.",
    "Workbook exceeds 100,000 populated cells.",
    "Workbook exceeds the 100,000 populated cell limit.",
    "Workbook history must be an array of at most 2,000 events.",
    "Workbook link cache exceeds its cell or sheet limit.",
    "Workbook link metadata exceeds 2 million characters.",
    "Workbook link targets must be unique.",
    "Workbook metadata requires a valid UTC timestamp.",
    "Workbook modified time cannot precede its creation time.",
    "Workbook name",
    "Workbook or authorization changed. Preview consolidation again.",
    "Workbook revision limit reached.",
    "Workbook structure is protected.",
    "Workbook structure is protected. Unprotect it before adding, deleting, renaming, moving or hiding worksheets.",
    "Workbook structure is protected. Unprotect it before changing names.",
    "Workbook structure is protected. Unprotect it before saving a new report.",
    "Workbook structure is protected. Unprotect the structure before changing worksheets.",
    "Workbook structure protection requires a valid salted password hash.",
    "Workbook styles must be an object.",
    "Workbook, extension or permissions changed. Preview again.",
    "Workbook, extension or permissions changed. Review inputs again.",
    "Workbook, extension or permissions changed. Run and review again.",
    "Workbook, source, review or permissions changed. Preview links again.",
    "Workbooks containing macros, ActiveX or embedded objects are not accepted. Save a clean XLSX copy first.",
    "Worker heap budget must be 16\u2013256 MiB.",
    "Worksheet no longer exists.",
    "Worksheet not found.",
    "Worksheet-qualified defined names are not resolved by this inspector.",
    "XLSX calculation properties cannot contain a DTD.",
    "XLSX column metadata must stay within 16,384 nonoverlapping columns.",
    "XLSX column widths must be finite nonnegative values.",
    "XLSX compatibility details exceed the 100,000-entry budget. Split the source workbook before importing.",
    "XLSX export exceeds the 10 MB reopen limit. Reduce the workbook or use native JSON.",
    "XLSX export exceeds the 100,000-hidden-row processing budget. Use native JSON to preserve larger hidden ranges without expanding them.",
    "XLSX files must be 10 MB or smaller.",
    "XLSX has more than 1,000 validation areas; narrow the source rules before import.",
    "XLSX hidden rows exceed the 100,000-row processing budget. Reduce hidden rows in the source; native JSON preserves larger visibility intervals.",
    "XLSX local headers do not match the archive directory.",
    "XLSX local sizes or checksum do not match the archive directory.",
    "XLSX page breaks exceed 100 boundaries per axis.",
    "XLSX validation ranges exceed the 100,000-cell import budget. Narrow the validation area in the source workbook.",
    "XLSX validation ranges exceed the 100,000-cell interchange budget. Use native JSON to preserve larger rules.",
    "XML cell text budget exceeded.",
    "XML document types and entities are not supported.",
    "XML needs one root element.",
    "XML nesting or node budget exceeded.",
    "XML table count exceeded.",
    "XML type prefix is undeclared.",
    "Zero rate and zero payment cannot determine a period count.",
    "actor",
    "amount",
    "approved range",
    "artifact ID",
    "blank-category",
    "bounds",
    "broken-link",
    "conflict",
    "corrupt",
    "cycle",
    "duplicate-category",
    "dynamic",
    "empty",
    "event ID",
    "excluded",
    "external",
    "external-link",
    "file name",
    "includeHidden must be boolean.",
    "includeNames must be boolean.",
    "invalid",
    "limit",
    "missing",
    "overlap",
    "reconciliation",
    "retainer",
    "run ID",
    "sheet ID",
    "source",
    "timestamp",
    "unavailable",
    "uncertain",
    "unsupported",
  ],
  k8 = [
    [
      "",
      " at ",
      "!",
      ", range ",
      ". Import stopped before changing the workbook. Open the original in Excel and convert the array to ordinary supported formulas or an explicitly reviewed values snapshot, then import that copy. Fixed CSE ranges and Excel array metadata are not silently treated as dynamic spills.",
    ],
    [
      "Duplicate/unique conditional rule at ",
      " cannot preserve Excel comparison semantics for this value. Use native JSON; XLSX supports finite numbers, canonical numeric text and short printable ASCII text without date-like numeric prefixes.",
    ],
    [
      "Broken pivot field \u201C",
      "\u201D: source field was not found. Restore the header or remove/reselect this field before refreshing. Historical totals are not current.",
    ],
    [
      "Workbook-scoped name ",
      " has ambiguous unqualified worksheet references. Qualify those references before inserting or deleting rows or columns.",
    ],
    [
      "Name ",
      " depends on worksheet scope. Keep the formula on its sheet or replace that name with explicit qualified references before moving it.",
    ],
    [
      "Reference ",
      "!",
      " includes cells outside the moved block. This move would require a split or cross-sheet range; use explicit references first.",
    ],
    [
      "Appended value in ",
      " does not match its ",
      " type. Supply a typed value or explicitly revise the column type; no rows were appended.",
    ],
    [
      "Source calculation error ",
      " at ",
      "!",
      ". Repair the source calculation before building or refreshing analysis; no totals were written.",
    ],
    [
      "This change would update dependent cells or analysis on protected worksheet ",
      ". Unprotect it before renaming the table field.",
    ],
    [
      "Calculated field \u201C",
      "\u201D divides by zero in a group or total. Filter, clean or change the calculation; no result was written.",
    ],
    [
      "Name ",
      " contains a current-row structured reference. Use an explicit qualified cell reference before moving this formula.",
    ],
    [
      "Calculation is incomplete at ",
      ". Recalculate before filtering; unchecked rows are never reported as an empty population.",
    ],
    [
      "This structural edit requires different relative ",
      " formulas within one rule. Split, relocate, or remove the rule first.",
    ],
    [
      "This selection is part of populated records. Expand the selection to ",
      ", or cancel; partial-record sorting is blocked.",
    ],
    [
      "Structural edits can verify at most 5,000 applied cells for relative ",
      " formulas. Narrow or remove the rules first.",
    ],
    [
      "The workbook-scoped definition ",
      " resolves differently across worksheets. Qualify its references before renaming.",
    ],
    [
      "Date grouping for \u201C",
      "\u201D needs YYYY-MM-DD text on source row ",
      "; choose the correct date system or clean the source.",
    ],
    [
      "Enter a valid YYYY-MM-DD date between ",
      " and 9999-12-31, or begin with an apostrophe to keep the entry as text.",
      "",
    ],
    [
      "",
      " can refer to cells dynamically. Replace dynamic references with explicit A1 references before moving cells.",
    ],
    [
      "Unsupported formula kind at ",
      " in ",
      ". Convert data-table formulas to reviewed ordinary formulas or values first.",
    ],
    [
      "This move would split ",
      " into separate ranges. Move the complete range or use explicit cell references first.",
    ],
    [
      "The move intersects drawing ",
      ". Drawing footprints cannot be moved or partially overwritten by this command.",
    ],
    [
      "Source calculation is incomplete (",
      "). Repair blocked or over-limit dependencies before building analysis.",
    ],
    [
      "An operational calculation error (",
      ") prevents duplicate review, even when a formula handler conceals it.",
    ],
    [
      "The source contains ",
      " hidden rows. Include hidden rows explicitly or cancel so records stay together.",
    ],
    [
      "Unsupported source token ",
      ". Templates, expressions and arbitrary JavaScript are outside this adapter.",
    ],
    [
      "Renaming would be shadowed by a local LET variable or LAMBDA parameter in ",
      ". Choose a distinct name.",
    ],
    [
      "XLSX checksum mismatch in ",
      ". The archive is corrupt; repair the original and save a clean XLSX copy.",
    ],
    [
      "Unsupported XLSX package type or Strict XML namespace in ",
      ". Save as standard Excel Workbook (.xlsx).",
    ],
    [
      "The move intersects table ",
      ". Table-cell moves are not supported; use ordinary cells outside tables.",
    ],
    [
      "Invalid row expression near character ",
      ". Use [Field], literals, operators and supported functions.",
    ],
    [
      "Inserting columns inside a ",
      " requires explicit field reconciliation. Resize or recreate it first.",
    ],
    [
      "Calculated field \u201C",
      "\u201D requires sum aggregation; its arithmetic is applied to source-field sums.",
    ],
    [
      "Corrupt XLSX compressed content in ",
      ". Repair the original workbook and save a clean XLSX copy.",
    ],
    [
      "Source ",
      " contains a formula. Copy its results to ordinary values explicitly before splitting.",
    ],
    [
      "Row ",
      " includes blank or nonnumeric values. Correct it or explicitly exclude incomplete rows.",
    ],
    [
      "Amount ",
      "!",
      " is included more than once. Remove overlapping sources to prevent double counting.",
    ],
    [
      "Ordered filter found a nonnumeric value in \u201C",
      "\u201D on input record ",
      ". Convert types explicitly.",
    ],
    [
      "Unsupported boolean value at ",
      " in ",
      "; save numeric Boolean encoding in a standard XLSX copy.",
    ],
    [
      "Group field ",
      " cannot filter measure ",
      ". Use dimensions shared by every selected fact table.",
    ],
    [
      "Structural edits cannot repair dynamic ",
      "-rule references. Remove the formula rule first.",
    ],
    [
      "This edit would remove a selected chart ",
      ". Reset the chart selection before deleting it.",
    ],
    [
      "Source ",
      " is missing or not approved. Attach the exact source and approve its worksheets.",
    ],
    [
      "Invalid ",
      " byte sequence. Choose the source encoding; bytes are never silently replaced.",
    ],
    [
      "Deleting columns inside a ",
      " would invalidate its fields. Resize or recreate it first.",
    ],
    [
      "Invalid ",
      " filter value. Use a finite number or a date in the workbook's date policy.",
    ],
    [
      "An assumption for ",
      "!",
      " violates its data validation constraints. Baseline unchanged.",
    ],
    [
      "This name is referenced by ",
      ". Remove or replace its dependencies before deletion.",
    ],
    [
      "Select at most ",
      " cells for this operation. The current selection contains ",
      " cells.",
    ],
    [
      "Unable to read delimited data: ",
      " Review the import schema; no rows were imported.",
    ],
    [
      "XLSX defined names: ",
      " Import/export stopped to prevent changed formula bindings.",
    ],
    [
      "Operational calculation error: ",
      ". An error handler cannot conceal this failure.",
    ],
    [
      "Source ",
      " ",
      " does not match ",
      ". No currency, unit or period conversion is implicit.",
    ],
    [
      "Column ",
      " contains mixed types. Clean the source before adding it to the model.",
    ],
    [
      "",
      " is hidden. Use View \u2192 Hide & Unhide or Data \u2192 Group & Outline to reveal it.",
    ],
    [
      "More than one filter path connects ",
      " to ",
      ". Remove the ambiguous relationship.",
    ],
    [
      "The related side of ",
      " is not unique. Choose one-to-many or correct its keys.",
    ],
    [
      "",
      " cannot be interpreted as a Number. Choose Automatic or correct the source.",
    ],
    [
      "Date grouping for \u201C",
      "\u201D requires a valid ",
      " numeric date serial on source row ",
      ".",
    ],
    [
      "This edit removes a ",
      " formula's anchor. Relocate or remove the rule first.",
    ],
    [
      "This edit would remove the anchor of ",
      ". Move or delete that drawing first.",
    ],
    [
      "Renaming would change an existing reference in ",
      ". Choose a distinct name.",
    ],
    [
      "This edit would remove a ",
      " header. Move or recreate that structure first.",
    ],
    [
      "CDATA content is unsupported in ",
      ". Resave as standard XLSX text encoding.",
    ],
    [
      "This worksheet already has the Excel limit of ",
      " rows and ",
      " columns (XFD).",
    ],
    [
      "This move would reorder part of ",
      ". Move the complete structure together.",
    ],
    [
      "Renaming would be shadowed by another name in ",
      ". Choose a distinct name.",
    ],
    [
      "Cell ",
      " belongs to the array formula in ",
      ". Edit the array formula instead.",
    ],
    [
      "Source ",
      " is a formula. Create a reviewed values snapshot before cleanup.",
    ],
    [
      "Declared dimension ",
      " exceeds or violates the supported worksheet bounds.",
    ],
    [
      "The XLSX exceeds ",
      " source cell records, including explicit blank cells.",
    ],
    [
      "Invalid picture ",
      ": use nonblank plain text without control characters.",
    ],
    [
      "",
      " must contain 1\u2013",
      " characters without unsupported control characters.",
    ],
    ["A recipe field exceeds the current ", " columns. Review steps in order."],
    ["", " uses dynamic or locally bound names and cannot be safely migrated."],
    [
      "Amount ",
      "!",
      " must be a finite number; text is not silently converted.",
    ],
    ["The structural edit would push a ", " beyond the worksheet boundary."],
    ["The resulting name has a cycle or exceeds the expansion limit: ", ""],
    ["Unmergeable array output at ", ". Move or remove its formula first."],
    ["Calculate ", " successfully before sorting. Its current result is ", "."],
    ["Source ", " is a formula. Create a reviewed values snapshot first."],
    ["Unsupported XML encoding in ", ". Save a standard UTF-8 XLSX copy."],
    [
      "Picture \u201C",
      "\u201D could not be decoded. Replace it before printing.",
    ],
    ["Reusable formula evaluation exceeds ", " levels or ", " work units."],
    [
      "Source field \u201C",
      "\u201D was not found. Refresh the field selection.",
    ],
    [
      "Source field \u201C",
      "\u201D was not found. Refresh the field selection.",
    ],
    ["What-if names must be nonempty text of at most ", " characters."],
    [
      "Picture \u201C",
      "\u201D could not be loaded. Replace it before printing.",
    ],
    ["Cell ", " has a duplicate, unordered or mismatched row address."],
    ["This edit removes a ", " formula range. Remove the rule first."],
    ["Cannot publish ", "!", ": resolve its error or array result first."],
    ["Custom formula must return TRUE/FALSE or a finite number", "."],
    [
      "Timeline field \u201C",
      "\u201D has an invalid ISO date on source row ",
      ".",
    ],
    ["What-if analysis requires supported static dependencies: ", ""],
    ["Cell ", " exceeds or violates the supported worksheet bounds."],
    ["The source schema of ", " changed. Review the model columns."],
    ["Date values must be whole Excel ", " serials between 0 and ", "."],
    ["", " must be a finite number in the selected numeric locale."],
    ["Date group \u201C", "\u201D must be selected as a row or column field."],
    [
      "Timeline field \u201C",
      "\u201D requires ",
      " date serials on source row ",
      ".",
    ],
    ["Required field \u201C", "\u201D is missing. Review the source schema."],
    ["", " must contain 1\u2013", " characters without control characters."],
    ["File BOM indicates ", "; select that encoding or Automatic."],
    ["Unsupported time zone: ", ". Use an IANA time zone or UTC."],
    ["", " is not a valid date under the workbook's date policy."],
    ["Worksheet ", ": ", " Import stopped; no cells were truncated."],
    ["Print titles support at most ", " ", " within the worksheet."],
    ["Each formula must return a finite number; received ", "."],
    ["Unknown identifier ", ". Column references use [Field]."],
    ["Control total for \u201C", "\u201D exceeds finite numeric limits."],
    ["Inspection limits must be integers between 1 and ", "."],
    ["Goal Seek requires supported static dependencies: ", ""],
    ["Calculated field \u201C", "\u201D exceeds finite numeric bounds."],
    ["Pivot controls require a list of at most ", " entries."],
    ["Workbook link review is limited to ", " formula cells."],
    ["Row heights must be ", "\u2013", " pixels for an existing row."],
    ["Derived field \u201C", "\u201D requires numbers on source row ", "."],
    ["Arithmetic paste requires numbers or blanks at ", "."],
    ["A transaction may contain at most ", " cell changes."],
    ["A transaction may contain at most ", " cell changes."],
    ["Replacement would exceed 32,767 characters at ", "!", "."],
    ["Derived field \u201C", "\u201D divides by zero on source row ", "."],
    ["Source worksheet ", " is outside the approved scope."],
    ["XML complexity exceeds the supported budget in ", "."],
    ["A workbook can contain at most ", " review threads."],
    ["Calculated field \u201C", "\u201D duplicates a source header."],
    ["Field \u201C", "\u201D is missing from the step input schema."],
    ["Unknown name or unsupported custom function: ", "."],
    ["Cell ", " exceeds the supported workbook bounds."],
    ["Sort field \u201C", "\u201D must be a row or column group."],
    ["DTD/entity declarations are unsupported in ", "."],
    ["External link ranges are limited to ", " cells."],
    ["Arithmetic paste would divide by zero at ", "."],
    ["Relationship ", " requires matching key types."],
    ["", ".", " has an invalid ", " value on worksheet row ", "."],
    ["The one side of ", " contains a duplicate key."],
    ["Measure ", " exceeds finite numeric precision."],
    ["Worksheet ", " is outside the approved sheet ", "."],
    ["Date serial is outside 0\u2013", " for the ", " system."],
    ["Relationship escapes the XLSX package in ", "."],
    ["Host selection may include at most ", " cells."],
    ["", " must be unique, including case variants."],
    ["Invalid ", ": use a name, not a cell address", "."],
    ["AutoFit ", " index is outside the worksheet."],
    ["The date is impossible or outside ", "\u20139999."],
    ["Select between 1 and ", " rectangular areas."],
    ["", " must be a bounded list of column names."],
    ["Malformed row: expected ", " fields, found ", "."],
    ["Repeated category also occurs at row ", ". ", ""],
    ["Invalid field transformation evidence: ", ""],
    ["LAMBDA expects ", " argument(s); received ", "."],
    ["Automation archive checksum mismatch: ", "."],
    ["Target ", " is outside the approved range ", "."],
    ["The one side of ", " contains a blank key."],
    ["Invalid ", "; limits can only be narrowed."],
    ["Unsupported Solver function or name: ", "."],
    ["Invalid or duplicate relationship in ", "."],
    ["Unsupported external relationship in ", "."],
    ["Model table ", " has a missing worksheet."],
    ["The structured table for ", " is missing."],
    ["", " must be text of at most ", " characters."],
    ["", " is outside the finite numeric range."],
    ["Worksheet ", " exceeds ", " rows or ", " columns."],
    ["Unsupported or invalid print object: ", ""],
    ["Arithmetic paste would overflow at ", "."],
    ["Invalid shared-string index at ", " in ", "."],
    ["The ", " distribution primitive failed."],
    ["Rule ", " has duplicate field mappings."],
    ["Invalid recipe execution evidence: ", ""],
    ["Malformed automation metadata in ", ": ", ""],
    ["Measure ", " requires a number column."],
    ["", " must be an integer between ", " and ", "."],
    ["Cell ", " is outside sheet dimensions."],
    ["Select at most ", " rectangular areas."],
    ["Invalid date in \u201C", "\u201D on source row ", "."],
    ["Unsupported cell style setting: ", "."],
    ["Invalid finite numeric value at ", "."],
    ["Unsafe internal relationship in ", "."],
    ["Missing relationship target ", " in ", "."],
    ["Invalid pattern-check settings: ", ""],
    ["Invalid consolidation evidence: ", ""],
    ["Unknown row-expression field [", "]."],
    ["Unsupported row function ", ". Use ", "."],
    ["What-if lists require ", "\u2013", " entries."],
    ["Delimited data exceeds ", " columns."],
    ["Invalid relationship target in ", "."],
    ["Unsupported error value at ", " in ", "."],
    ["Paste at most ", " cells at a time."],
    ["", " must be an integer from 1 to ", "."],
    ["Invalid or out-of-bounds cell: ", ""],
    ["Unsupported delimited ", " setting."],
    ["", " must be an ISO UTC timestamp."],
    ["Invalid field transformation: ", ""],
    ["Invalid numeric value at ", " in ", "."],
    ["Source worksheet ", " is missing."],
    ["Corrupt automation archive: ", "."],
    ["Unsupported XML encoding in ", "."],
    ["Delimited data exceeds ", " rows."],
    ["Invalid cell style index in ", "."],
    ["Invalid analysis request: ", ": ", ""],
    ["", " must contain 1\u2013", " characters."],
    ["", " must contain 1\u2013", " characters."],
    ["Row ", " must be between 1 and ", "."],
    ["Cell calculation returned ", "."],
    ["", " arguments cannot be empty."],
    ["", " contains duplicate fields."],
    ["Calculated column formula: ", ""],
    ["Unsupported cell type ", " in ", "."],
    ["", " must be a finite number."],
    ["Invalid analysis report: ", ""],
    ["", " must be a finite number."],
    ["", " must be a plain object."],
    ["Circular defined name: ", "."],
    ["Invalid export ", " control."],
    ["Style ", " must be boolean."],
    ["Invalid consolidation: ", ""],
    ["Unknown measure table ", "."],
    ["Invalid saved recipes: ", ""],
    ["Invalid data model: ", ": ", ""],
    ["Unknown model field ", ".", "."],
    ["Unknown query field ", ".", "."],
    ["Invalid model report: ", ""],
    ["", " requires ", "\u2013", " arguments."],
    ["Invalid cell address: ", ""],
    ["Invalid model query: ", ""],
    ["Duplicate ", " binding: ", "."],
    ["Unsupported ", " setting."],
    ["Invalid automation ", "."],
    ["Invalid validation ", "."],
    ["", " expects ", " arguments."],
    ["Missing XLSX part ", "."],
    ["Formula returned ", "."],
    ["Input record ", " (", "): ", ""],
    ["Input record ", " (", "): ", ""],
    ["Malformed XML in ", "."],
    ["Invalid drawing ", "."],
    ["Invalid drawing ", "."],
    ["", " must be boolean."],
    ["Sheet not found: ", ""],
    ["", " print issue(s). ", ""],
    ["Invalid recipe: ", ""],
    ["Invalid chart ", "."],
    ["Invalid range: ", ""],
  ];
var Mk = Object.freeze({
    FILE_LIMIT:
      "File processing exceeds the supported size, range, nesting or record budget. Reduce the source and retry.",
    FILE_TYPE:
      "Choose a supported file with a safe filename and the correct extension.",
    FILE_INVALID:
      "The file contains invalid or unsupported records. Repair it in the source application and retry; the open workbook is unchanged.",
    FILE_ACTIVE:
      "Active or unsafe file content is not supported. Remove macros, embedded objects, unsafe links or document types and save a clean copy.",
    FILE_TIMEOUT:
      "File processing exceeded its time budget and was stopped. Reduce the source and retry; the open workbook is unchanged.",
    FILE_CANCELLED:
      "File processing was cancelled. The open workbook is unchanged.",
    FILE_WORKER:
      "File processing could not start or its worker stopped. Allow the trusted file worker in the host content security policy, or supply a worker factory, then retry.",
    FILE_BUSY:
      "Another file is being processed. Wait for it to finish or cancel it, then retry.",
    FILE_READ:
      "The file could not be read completely. Select the original file again; the open workbook is unchanged.",
  }),
  An = class extends Error {
    constructor(n, i = Mk[n]) {
      super(i);
      this.code = n;
      this.name = "FileProcessingError";
    }
  },
  Bq = new Set([
    ...x8,
    "Legacy CSE array formula",
    "Metadata-qualified Excel array formula (dynamic/legacy classification requires Excel metadata)",
  ]),
  Ok =
    " Open the original in Excel, repair it if needed, and save a clean, unencrypted XLSX copy.",
  jq = k8.map((e) => ({
    parts: e,
    expression: new RegExp(
      "^" +
        e
          .map((t) => t.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"))
          .join("([\\s\\S]*?)") +
        "$",
    ),
  }));
function Sg(e, t = 0) {
  if (!(e.length > 4096 || t > 3)) {
    if (Bq.has(e)) return e;
    if (e.endsWith(Ok)) {
      let n = Sg(e.slice(0, -Ok.length), t + 1);
      if (n) return n + Ok;
    }
    for (let { parts, expression } of jq) {
      let r = expression.exec(e);
      if (r)
        return parts
          .map(
            (o, a) =>
              o +
              (a < parts.length - 1
                ? (Sg(r[a + 1], t + 1) ?? "[source detail omitted]")
                : ""),
          )
          .join("");
    }
  }
}
function Wq(e) {
  if (e instanceof An) return new An(e.code, Sg(e.message) ?? Mk[e.code]);
  let t = e instanceof Error ? e.message : "",
    n =
      /exceed|budget|limit|too (?:large|many)|at most|smaller|depth|nesting/i.test(
        t,
      )
        ? "FILE_LIMIT"
        : /macro|active.?content|activex|executable|unsafe|document type|doctype|entity|dangerous|DDE/i.test(
              t,
            )
          ? "FILE_ACTIVE"
          : "FILE_INVALID";
  return new An(n, Sg(t) ?? Mk[n]);
}
function _8(e) {
  let { code, message } = Wq(e);
  return { code, message };
}
var Nc = Object.freeze({
  maxFileBytes: Ne.maxFileBytes,
  maxNativeBytes: Ne.maxNativeBytes,
  maxConfigurationBytes: 2e6,
  maxPictureBytes: 2 * 1024 * 1024,
  maxConcurrent: 2,
  defaultTimeoutMs: 3e4,
  maxTimeoutMs: 6e4,
  maxJsonDepth: 128,
  maxJsonNodes: 2e6,
  nodeHeapMiB: 256,
  maxResultBytes: 100 * 1024 * 1024,
});
function $k(e, t) {
  return e === "source"
    ? 4 * 1024 * 1024
    : e === "json"
      ? Nc.maxConfigurationBytes
      : e === "picture"
        ? Nc.maxPictureBytes
        : e === "workbook" && /\.json$/i.test(t)
          ? Ne.maxNativeBytes
          : Ne.maxFileBytes;
}
function E8(e, t, n) {
  if (
    typeof t != "string" ||
    !t.length ||
    t.length > 255 ||
    /[\\/\x00-\x1f\x7f\u202a-\u202e\u2066-\u2069]/.test(t)
  )
    throw new An("FILE_TYPE");
  if (e === "workbook" && /\.(xlsm|xlsb|xls|xltm|xlam)$/i.test(t))
    throw new An(
      "FILE_TYPE",
      "Legacy or macro-enabled Excel formats are not supported. Save as XLSX, CSV or native JSON.",
    );
  if (
    !{
      workbook: /\.(xlsx|xltx|csv|tsv|json)$/i,
      source: /\.(json|xml|pdf|png|jpe?g|html?)$/i,
      formats: /\.(xls|xlsb|xlsm|xltx|xltm|ods)$/i,
      delimited: /\.(csv|tsv)$/i,
      automation: /\.(bas|vba|cls|ts|osts|xlsm|xlam|xltm|xlsx|xltx|xls|xlsb)$/i,
      json: /\.json$/i,
      picture: /\.(png|jpe?g)$/i,
    }[e]?.test(e === "picture" ? t.trim() : t)
  )
    throw new An("FILE_TYPE");
  if (!Number.isSafeInteger(n) || n < 0 || n > $k(e, t))
    throw new An(
      "FILE_LIMIT",
      $k(e, t) === Ne.maxNativeBytes
        ? "Native workbook exceeds the 50 MB import limit."
        : $k(e, t) === Ne.maxFileBytes
          ? "Files must be 10 MB or smaller."
          : undefined,
    );
}
for (let e of [
  "fetch",
  "XMLHttpRequest",
  "WebSocket",
  "EventSource",
  "importScripts",
  "Worker",
  "SharedWorker",
  "BroadcastChannel",
  "indexedDB",
  "caches",
])
  Object.defineProperty(globalThis, e, {
    value: undefined,
    writable: false,
    configurable: false,
  });
function S8(e) {
  let t = JSON.parse(new TextDecoder("utf-8", { fatal: true }).decode(e)),
    n = [{ value: t, depth: 0 }],
    i = 0;
  for (; n.length;) {
    let r = n.pop();
    if (++i > Nc.maxJsonNodes || r.depth > Nc.maxJsonDepth)
      throw new An("FILE_LIMIT");
    if (r.value && typeof r.value == "object")
      for (let [o, a] of Object.entries(r.value)) {
        if (["__proto__", "prototype", "constructor"].includes(o))
          throw new An("FILE_ACTIVE");
        n.push({ value: a, depth: r.depth + 1 });
      }
  }
  return t;
}
var Pk = self,
  C8 = false;
Pk.onmessage = async ({ data }) => {
  if (C8) return;
  C8 = true;
  let t = data?.id;
  try {
    if (
      typeof t != "string" ||
      t.length > 100 ||
      !(data.bytes instanceof Uint8Array)
    )
      throw new An("FILE_INVALID");
    E8(data.operation, data.fileName, data.bytes.length);
    let n;
    switch (data.operation) {
      case "source":
        n = gk(data.bytes, data.fileName);
        break;
      case "formats":
        n = Vd(data.bytes, data.fileName);
        break;
      case "workbook": {
        /\.json$/i.test(data.fileName) && S8(data.bytes);
        let i = data.bytes,
          r = await kg({
            name: data.fileName,
            size: i.length,
            arrayBuffer: async () => i.buffer,
            text: async () =>
              new TextDecoder("utf-8", { fatal: true }).decode(i),
          });
        if (
          (/\.json$/i.test(data.fileName) &&
            r.workbook.xlsxSource &&
            w8(r.workbook),
          /\.json$/i.test(data.fileName) && r.workbook.alternateSource)
        ) {
          let o = r.workbook.alternateSource;
          Vd(lg(o), o.fileName);
        }
        if (/\.json$/i.test(data.fileName))
          for (let o of r.workbook.sourceImports ?? []) {
            let a = gk(ck(o.original), o.original.fileName);
            if (
              a.method !== "document-adapter" &&
              JSON.stringify(a.tables.find((s) => s.id === o.table.id)) !==
                JSON.stringify(o.table)
            )
              throw Error(
                "Retained structured evidence does not match its original.",
              );
          }
        n = r;
        break;
      }
      case "delimited": {
        let i = b8(data.bytes, data.fileName, data.options);
        n = {
          report: i,
          ...(i.valid
            ? { workbook: Eg(data.bytes, data.fileName, data.options).workbook }
            : {}),
        };
        break;
      }
      case "automation":
        n = o8(data.bytes, data.fileName);
        break;
      case "json":
        n = S8(data.bytes);
        break;
      case "picture": {
        let i = "";
        for (let a of data.bytes) i += String.fromCharCode(a);
        let r = `data:image/${/\.png$/i.test(data.fileName.trim()) ? "png" : "jpeg"};base64,${btoa(i)}`,
          o = Ha(r);
        n = { dataURL: r, width: o.width, height: o.height };
        break;
      }
      default:
        throw new An("FILE_TYPE");
    }
    if (new TextEncoder().encode(JSON.stringify(n)).length > Nc.maxResultBytes)
      throw new An("FILE_LIMIT");
    Pk.postMessage({ id: t, ok: true, value: n });
  } catch (n) {
    Pk.postMessage({ id: t, ok: false, error: _8(n) });
  }
};
