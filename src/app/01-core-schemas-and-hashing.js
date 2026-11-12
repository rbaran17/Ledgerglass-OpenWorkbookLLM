import ne from "react";
import {
  iso,
  string,
  union,
  number,
  boolean,
  null as _null,
  array,
  strictObject,
  enum as _enum,
  discriminatedUnion,
  literal,
  record,
  unknown,
} from "zod";

var objectDefineProperty = Object.defineProperty;
var defineDataProperty = (e, r, t) =>
  r in e
    ? objectDefineProperty(e, r, {
        enumerable: true,
        configurable: true,
        writable: true,
        value: t,
      })
    : (e[r] = t);
var defineDataPropertyByKey = (e, r, t) =>
  defineDataProperty(e, typeof r != "symbol" ? r + "" : r, t);
function copyModuleProperties(e, r) {
  for (var t = 0; t < r.length; t++) {
    const n = r[t];
    if (typeof n != "string" && !Array.isArray(n)) {
      for (const o in n)
        if (o !== "default" && !(o in e)) {
          const a = Object.getOwnPropertyDescriptor(n, o);
          a &&
            Object.defineProperty(
              e,
              o,
              a.get ? a : { enumerable: true, get: () => n[o] },
            );
        }
    }
  }
  return Object.freeze(
    Object.defineProperty(e, Symbol.toStringTag, { value: "Module" }),
  );
}
function unwrapDefaultExport(e) {
  return e && e.__esModule && Object.prototype.hasOwnProperty.call(e, "default")
    ? e.default
    : e;
}

const Gq = unwrapDefaultExport(ne),
  uEe = copyModuleProperties({ __proto__: null, default: Gq }, [ne]),
  deepFrozenValueSet = new WeakSet();
function isDeepFrozenValue(e) {
  return !!e && typeof e == "object" && deepFrozenValueSet.has(e);
}
function deepFreezeValue(e) {
  if (!e || typeof e != "object" || deepFrozenValueSet.has(e)) return e;
  for (const r of Object.values(e)) deepFreezeValue(r);
  return (Object.freeze(e), deepFrozenValueSet.add(e), e);
}
const linkedOrBoundCellCache = new WeakMap();
function hasLinkedOrBoundCells(e) {
  if (!isDeepFrozenValue(e)) return true;
  let r = linkedOrBoundCellCache.get(e);
  if (r === undefined) {
    r = false;
    for (const n in e)
      if (e[n].hyperlink || e[n].style?.binding) {
        r = true;
        break;
      }
    linkedOrBoundCellCache.set(e, r);
  }
  return r;
}
const serializedByteCache = new WeakMap(),
  utf8Encoder = new TextEncoder();
function measureJsonStringBytes(e) {
  return /^[\x20-\x21\x23-\x5b\x5d-\x7e]*$/.test(e)
    ? e.length + 2
    : utf8Encoder.encode(JSON.stringify(e)).byteLength;
}
function measureSerializedBytes(e) {
  if (typeof e == "string") return measureJsonStringBytes(e);
  if (!e || typeof e != "object") return JSON.stringify(e)?.length ?? 0;
  const r = isDeepFrozenValue(e) ? serializedByteCache.get(e) : undefined;
  if (r !== undefined) return r;
  let t = 2,
    n = 0;
  if (Array.isArray(e))
    for (const a of e)
      t += (n++ ? 1 : 0) + measureSerializedBytes(a === undefined ? null : a);
  else
    for (const a in e) {
      if (!Object.prototype.hasOwnProperty.call(e, a)) continue;
      const s = e[a];
      s !== undefined &&
        (t +=
          (n++ ? 1 : 0) +
          measureJsonStringBytes(a) +
          1 +
          measureSerializedBytes(s));
    }
  return (isDeepFrozenValue(e) && serializedByteCache.set(e, t), t);
}
const isoDateTimeSchemaFactory = iso.datetime;
function sha256HexDigest(e) {
  if (
    !ArrayBuffer.isView(e) ||
    Object.prototype.toString.call(e) !== "[object Uint8Array]" ||
    e.length > 50 * 1024 * 1024
  )
    throw Error("Hashing requires at most 50 MB of bytes.");
  const r = [
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
    t = [
      1779033703, 3144134277, 1013904242, 2773480762, 1359893119, 2600822924,
      528734635, 1541459225,
    ],
    n = new Uint8Array(Math.ceil((e.length + 9) / 64) * 64);
  (n.set(e), (n[e.length] = 128));
  const o = new DataView(n.buffer);
  o.setUint32(n.length - 4, e.length * 8);
  const a = new Uint32Array(64),
    s = (d, u) => (d >>> u) | (d << (32 - u));
  for (let d = 0; d < n.length; d += 64) {
    for (let g = 0; g < 16; g++) a[g] = o.getUint32(d + g * 4);
    for (let g = 16; g < 64; g++) {
      const b = a[g - 15],
        y = a[g - 2];
      a[g] =
        (a[g - 16] +
          (s(b, 7) ^ s(b, 18) ^ (b >>> 3)) +
          a[g - 7] +
          (s(y, 17) ^ s(y, 19) ^ (y >>> 10))) >>>
        0;
    }
    let [u, l, c, h, f, p, m, v] = t;
    for (let g = 0; g < 64; g++) {
      const b =
          (v +
            (s(f, 6) ^ s(f, 11) ^ s(f, 25)) +
            ((f & p) ^ (~f & m)) +
            r[g] +
            a[g]) >>>
          0,
        y =
          ((s(u, 2) ^ s(u, 13) ^ s(u, 22)) + ((u & l) ^ (u & c) ^ (l & c))) >>>
          0;
      ((v = m),
        (m = p),
        (p = f),
        (f = (h + b) >>> 0),
        (h = c),
        (c = l),
        (l = u),
        (u = (b + y) >>> 0));
    }
    [u, l, c, h, f, p, m, v].forEach((g, b) => (t[b] = (t[b] + g) >>> 0));
  }
  return t.map((d) => d.toString(16).padStart(8, "0")).join("");
}
const transformationRecipeLimits = Object.freeze({
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
  recipeStoreFormatId = "ledgerglass-recipes-v1",
  singleLineTextSchema = string()
    .trim()
    .min(1)
    .max(128)
    .refine((e) => !/[\u0000-\u001f\u007f]/.test(e)),
  opaqueIdSchema = string().min(1).max(100),
  cellScalarSchema = union([
    string().max(32767),
    number().finite(),
    boolean(),
    _null(),
  ]),
  distinctFieldListSchema = array(singleLineTextSchema)
    .max(50)
    .refine((e) => new Set(e).size === e.length, "Fields must be distinct."),
  sheetRangeSourceSchema = strictObject({
    sheetId: opaqueIdSchema,
    range: string().min(1).max(50),
    tableId: opaqueIdSchema.optional(),
  }),
  aggregateMeasureSchema = strictObject({
    field: singleLineTextSchema,
    operation: _enum([
      "sum",
      "count",
      "average",
      "min",
      "max",
      "distinctCount",
    ]),
    name: singleLineTextSchema,
  }),
  filterConditionSchema = strictObject({
    field: singleLineTextSchema,
    operator: _enum([
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
    value: cellScalarSchema.optional(),
  }),
  od = { id: opaqueIdSchema },
  transformationStepSchema = discriminatedUnion("kind", [
    strictObject({
      ...od,
      kind: literal("select"),
      fields: distinctFieldListSchema.min(1),
    }),
    strictObject({
      ...od,
      kind: literal("rename"),
      mappings: array(
        strictObject({ from: singleLineTextSchema, to: singleLineTextSchema }),
      )
        .min(1)
        .max(50),
    }),
    strictObject({
      ...od,
      kind: literal("filter"),
      conditions: array(filterConditionSchema).min(1).max(20),
      join: _enum(["and", "or"]),
    }),
    strictObject({
      ...od,
      kind: literal("types"),
      fields: distinctFieldListSchema.min(1),
      to: _enum(["text", "number", "integer", "boolean", "date"]),
      decimal: _enum([".", ","]),
      group: _enum(["", ".", ",", " ", "'"]),
      dateOrder: _enum(["MDY", "DMY"]),
    }),
    strictObject({
      ...od,
      kind: literal("replace"),
      fields: distinctFieldListSchema.min(1),
      mode: _enum(["value", "text"]),
      find: cellScalarSchema,
      replacement: cellScalarSchema,
    }),
    strictObject({
      ...od,
      kind: literal("derive"),
      name: singleLineTextSchema,
      expression: string().min(1).max(1e3),
    }),
    strictObject({
      ...od,
      kind: literal("sort"),
      fields: array(
        strictObject({
          field: singleLineTextSchema,
          direction: _enum(["ascending", "descending"]),
        }),
      )
        .min(1)
        .max(8),
    }),
    strictObject({
      ...od,
      kind: literal("unpivot"),
      fields: distinctFieldListSchema.min(1),
      attribute: singleLineTextSchema,
      value: singleLineTextSchema,
      keepNulls: boolean(),
    }),
    strictObject({
      ...od,
      kind: literal("group"),
      keys: distinctFieldListSchema,
      values: array(aggregateMeasureSchema).min(1).max(20),
    }),
    strictObject({
      ...od,
      kind: literal("pivot"),
      keys: distinctFieldListSchema,
      column: singleLineTextSchema,
      value: singleLineTextSchema,
      operation: aggregateMeasureSchema.shape.operation,
    }),
    strictObject({
      ...od,
      kind: literal("append"),
      source: sheetRangeSourceSchema,
    }),
    strictObject({
      ...od,
      kind: literal("merge"),
      source: sheetRangeSourceSchema,
      mode: _enum(["inner", "left", "right", "full", "leftAnti", "rightAnti"]),
      keys: array(
        strictObject({
          left: singleLineTextSchema,
          right: singleLineTextSchema,
        }),
      )
        .min(1)
        .max(8),
      prefix: singleLineTextSchema,
      normalization: strictObject({
        trim: boolean().optional(),
        collapseWhitespace: boolean().optional(),
        ignoreCase: boolean().optional(),
        unicode: _enum(["NFC", "NFKC"]).optional(),
      }).optional(),
      fuzzy: strictObject({
        key: number().int().min(0).max(7),
        threshold: number().finite().min(0).max(1),
      }).optional(),
    }),
  ]),
  transformationRecipeSchema = strictObject({
    engine: literal(recipeStoreFormatId),
    source: sheetRangeSourceSchema,
    steps: array(transformationStepSchema).max(
      transformationRecipeLimits.steps,
    ),
  }).superRefine((e, r) => {
    (new Set(e.steps.map((t) => t.id)).size !== e.steps.length &&
      r.addIssue({ code: "custom", message: "Step IDs must be unique." }),
      new Set(
        [
          e.source,
          ...e.steps.flatMap((t) =>
            t.kind === "append" || t.kind === "merge" ? [t.source] : [],
          ),
        ].map((t) => JSON.stringify(t)),
      ).size > transformationRecipeLimits.sources &&
        r.addIssue({
          code: "custom",
          message: "Use no more than eight recipe sources.",
        }));
  });
function canonicalJsonStringify(e) {
  return JSON.stringify(e, (r, t) =>
    t && typeof t == "object" && !Array.isArray(t)
      ? Object.fromEntries(
          Object.keys(t)
            .sort()
            .map((n) => [n, t[n]]),
        )
      : t,
  );
}
function fingerprintCanonicalValue(e) {
  const r = canonicalJsonStringify(e);
  if (r.length > 16e6)
    throw Error(
      "Recipe fingerprint exceeds the 16-million-character evidence budget.",
    );
  return sha256HexDigest(new Uint8Array(new TextEncoder().encode(r)));
}
function parseTransformationRecipe(e) {
  const r = transformationRecipeSchema.safeParse(e);
  if (!r.success) throw Error(`Invalid recipe: ${r.error.issues[0]?.message}`);
  return r.data;
}
const sha256HexSchema = string().regex(/^[0-9a-f]{64}$/),
  nonnegativeSafeIntegerSchema = number()
    .int()
    .nonnegative()
    .max(Number.MAX_SAFE_INTEGER - 1),
  isoTimestampSchema = string().refine(
    (e) => Number.isFinite(Date.parse(e)) && new Date(e).toISOString() === e,
  ),
  transformationRecipeVersionSchema = strictObject({
    version: number().int().min(1).max(transformationRecipeLimits.versions),
    definition: transformationRecipeSchema,
    definitionHash: sha256HexSchema,
    createdAt: isoTimestampSchema,
  }),
  SJ = array(
    strictObject({
      id: opaqueIdSchema,
      name: singleLineTextSchema,
      versions: array(transformationRecipeVersionSchema)
        .min(1)
        .max(transformationRecipeLimits.versions),
    }),
  ).max(transformationRecipeLimits.recipes),
  columnTypeProfileSchema = strictObject({
    name: singleLineTextSchema,
    types: array(_enum(["blank", "number", "string", "boolean"])).max(4),
    declaredType: string().max(32).optional(),
  }),
  numericColumnSummarySchema = strictObject({
    field: singleLineTextSchema,
    numericCount: nonnegativeSafeIntegerSchema,
    blankCount: nonnegativeSafeIntegerSchema,
    sum: number().finite(),
  }),
  transformationStepRunSchema = strictObject({
    index: number().int().min(1).max(20),
    id: opaqueIdSchema,
    kind: string().max(32),
    status: _enum(["complete", "failed", "skipped"]),
    inputSchema: array(columnTypeProfileSchema).max(50),
    inputRows: nonnegativeSafeIntegerSchema,
    outputSchema: array(columnTypeProfileSchema).max(50).optional(),
    outputRows: nonnegativeSafeIntegerSchema.optional(),
    inputHash: sha256HexSchema.optional(),
    outputHash: sha256HexSchema.optional(),
    controls: array(numericColumnSummarySchema).max(50).optional(),
    error: string().max(4e3).optional(),
    details: record(string(), unknown()).optional(),
  }),
  transformationRunEvidenceSchema = strictObject({
    engine: literal(recipeStoreFormatId),
    recipeId: opaqueIdSchema,
    recipeName: singleLineTextSchema,
    version: number().int().min(1).max(20),
    definition: transformationRecipeSchema,
    definitionHash: sha256HexSchema,
    sourceWorkbookId: opaqueIdSchema,
    sourceRevision: nonnegativeSafeIntegerSchema,
    sourceHash: sha256HexSchema,
    outputHash: sha256HexSchema,
    outputRows: number().int().min(0).max(1e3),
    outputRange: string().min(1).max(50),
    at: isoTimestampSchema,
    steps: array(transformationStepRunSchema).max(20),
    controls: array(numericColumnSummarySchema).max(50),
    lineage: record(
      string().regex(/^[1-9]\d*$/),
      array(string().min(1).max(200)).max(4e3),
    ),
  });
function validateSavedRecipeMetadata(e, r) {
  const t = e;
  if (t.transformationRecipes !== undefined) {
    if (
      JSON.stringify(t.transformationRecipes).length >
      transformationRecipeLimits.metadataCharacters
    )
      throw Error("Saved recipe metadata exceeds 2,000,000 characters.");
    const d = SJ.safeParse(t.transformationRecipes);
    if (!d.success)
      throw Error(`Invalid saved recipes: ${d.error.issues[0]?.message}`);
    if (
      new Set(d.data.map((u) => u.id)).size !== d.data.length ||
      new Set(d.data.map((u) => u.name.toLowerCase())).size !== d.data.length
    )
      throw Error("Saved recipe names and IDs must be unique.");
    for (const u of d.data)
      for (const [l, c] of u.versions.entries())
        if (
          c.version !== l + 1 ||
          c.definitionHash !== fingerprintCanonicalValue(c.definition)
        )
          throw Error(
            "Saved recipe versions or definition hashes are inconsistent.",
          );
    r.transformationRecipes = d.data;
  }
  let n = 0;
  for (const d of r.sheets) {
    const u = t.sheets?.find((h) => h.id === d.id);
    if (u?.transformationRun === undefined) continue;
    if (
      ((n += JSON.stringify(u.transformationRun).length),
      n > transformationRecipeLimits.runMetadataCharacters)
    )
      throw Error(
        "Recipe execution evidence exceeds 8,000,000 characters. Retain a native backup and remove obsolete output sheets.",
      );
    const l = transformationRunEvidenceSchema.safeParse(u.transformationRun);
    if (!l.success)
      throw Error(
        `Invalid recipe execution evidence: ${l.error.issues[0]?.message}`,
      );
    const c = l.data;
    if (
      c.definitionHash !== fingerprintCanonicalValue(c.definition) ||
      c.steps.length !== c.definition.steps.length ||
      c.steps.some(
        (h, f) =>
          h.status !== "complete" ||
          h.index !== f + 1 ||
          h.id !== c.definition.steps[f].id ||
          h.kind !== c.definition.steps[f].kind,
      ) ||
      Object.values(c.lineage).reduce((h, f) => h + f.length, 0) >
        transformationRecipeLimits.lineageLinks ||
      Object.keys(c.lineage).some(
        (h) => Number(h) < 2 || Number(h) > c.outputRows + 1,
      ) ||
      Object.keys(c.lineage).length !== c.outputRows
    )
      throw Error("Recipe execution metadata is inconsistent.");
    d.transformationRun = c;
  }
}
const normalizeFunctionName = (e) =>
    e.replace(/^_xlfn\.(?:_xlws\.)?/i, "").toUpperCase(),
  normalizeBindingName = (e) => e.replace(/^_xlpm\./i, "").toLowerCase();
function isValidFormulaName(e, r = false) {
  return (
    (e = e.replace(/^_xlpm\./i, "")),
    /^[A-Za-z_\\][A-Za-z0-9_.\\]{0,254}$/.test(e) &&
      !/^(?:[A-Za-z]{1,3}[1-9]\d*|R\d+C\d+|R|C)$/i.test(e) &&
      !/^__LG/i.test(e) &&
      (!r || !e.includes("."))
  );
}
function tokenizeFormulaText(e) {
  let r = e.startsWith("=") ? 1 : 0;
  const t = () => {
      throw Error("Unbalanced or malformed formula syntax.");
    },
    n = (d) => {
      for (r++; r < e.length;)
        if (e[r++] === d)
          if (e[r] === d) r++;
          else return;
      t();
    },
    o = (d, u) => {
      let l = 0;
      do {
        const c = e[r];
        if (c === '"' || (c === "'" && d !== "[")) {
          n(c);
          continue;
        }
        (c === d && l++, c === u && l--, r++);
      } while (r < e.length && l);
      l && t();
    },
    a = (d) => {
      if (d > 128) throw Error("#LIMIT!");
      const u = [];
      for (; r < e.length && ![",", ")"].includes(e[r]);) {
        const c = r;
        let h;
        const f = e[r];
        if (f === '"') n(f);
        else if (f === "{") o("{", "}");
        else if (f === "[") o("[", "]");
        else if (f === "(") {
          r++;
          const p = a(d + 1);
          (e[r++] !== ")" && t(),
            (h = { kind: "group", body: p, start: c, end: r }));
        } else {
          let p;
          const m =
            /^(?:'((?:[^']|'')+)'|([A-Za-z0-9_\\][A-Za-z0-9_.\\]*))!/.exec(
              e.slice(r),
            );
          m && ((p = m[1]?.replace(/''/g, "'") ?? m[2]), (r += m[0].length));
          const v =
            /^(?:#(?:REF!|NAME\?|VALUE!|DIV\/0!|N\/A|NUM!|NULL!|SPILL!|CALC!|CYCLE!|LIMIT!|BLOCKED!|ERROR!)|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}(?![\w.])|\$?[1-9]\d*:\$?[1-9]\d*(?![\w.])|\$?[A-Za-z]{1,3}\$?[1-9]\d*(?![\w.\\])|(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?)/i.exec(
              e.slice(r),
            );
          if (v) r += v[0].length;
          else {
            const g = /^[A-Za-z_\\][A-Za-z0-9_.\\]*/.exec(e.slice(r));
            g
              ? ((r += g[0].length),
                e[r] === "["
                  ? o("[", "]")
                  : (h = {
                      kind: "symbol",
                      name: g[0],
                      qualifier: p,
                      start: c,
                      end: r,
                    }))
              : ((p || f === "'" || f === "}") && t(), r++);
          }
        }
        if (
          ((h ??= { kind: "text", text: e.slice(c, r), start: c, end: r }),
          h.kind !== "text")
        )
          for (;;) {
            const p = /^\s*\(/.exec(e.slice(r));
            if (!p) break;
            r += p[0].length;
            const m = [];
            if (e[r] !== ")") for (; m.push(a(d + 1)), e[r] === ",";) r++;
            (e[r++] !== ")" && t(),
              (h = { kind: "call", callee: h, args: m, start: c, end: r }));
          }
        u.push(h);
      }
      return u;
    },
    s = a(0);
  return (r !== e.length && t(), s);
}
const dropBlankTextTokens = (e) =>
  e.filter((r) => r.kind !== "text" || r.text.trim());
function validateBindingName(e, r = false) {
  const t = dropBlankTextTokens(e),
    n = t[0];
  if (
    t.length !== 1 ||
    n?.kind !== "symbol" ||
    n.qualifier ||
    !isValidFormulaName(n.name, r)
  )
    throw Error(
      `Invalid ${r ? "LAMBDA parameter" : "LET variable"}: use a name, not a cell address${r ? ", without a period" : ""}.`,
    );
  return n;
}
function validateLetLambdaCall(e) {
  if (e.callee.kind !== "symbol" || e.callee.qualifier) return;
  const r = normalizeFunctionName(e.callee.name);
  if (r !== "LET" && r !== "LAMBDA") return;
  if (
    r === "LET" &&
    (e.args.length < 3 || e.args.length > 253 || e.args.length % 2 !== 1)
  )
    throw Error("LET requires 1–126 name/value pairs and a final calculation.");
  if (r === "LAMBDA" && (e.args.length < 1 || e.args.length > 254))
    throw Error("LAMBDA requires a calculation and at most 253 parameters.");
  if (e.args.some((n) => !dropBlankTextTokens(n).length))
    throw Error(`${r} arguments cannot be empty.`);
  const t = new Set();
  for (let n = 0; n < e.args.length - 1; n += r === "LET" ? 2 : 1) {
    const o = normalizeBindingName(
      validateBindingName(e.args[n], r === "LAMBDA").name,
    );
    if (t.has(o)) throw Error(`Duplicate ${r} binding: ${o}.`);
    t.add(o);
  }
  return r;
}
function walkFormulaNodes(e, r, t = new Set()) {
  for (const n of e) {
    if (
      (n.kind === "symbol" && r(n, t, false, false),
      n.kind === "group" && walkFormulaNodes(n.body, r, t),
      n.kind !== "call")
    )
      continue;
    const o = validateLetLambdaCall(n);
    if (
      (n.callee.kind === "symbol"
        ? r(n.callee, t, true, false)
        : walkFormulaNodes([n.callee], r, t),
      !o)
    ) {
      n.args.forEach((s) => walkFormulaNodes(s, r, t));
      continue;
    }
    const a = new Set(t);
    for (let s = 0; s < n.args.length - 1; s += o === "LET" ? 2 : 1) {
      const d = validateBindingName(n.args[s], o === "LAMBDA");
      (r(d, a, false, true),
        o === "LET" && walkFormulaNodes(n.args[s + 1], r, a),
        a.add(normalizeBindingName(d.name)));
    }
    walkFormulaNodes(n.args.at(-1), r, a);
  }
}
function parseFormulaTree(e) {
  try {
    const r = dropBlankTextTokens(tokenizeFormulaText(e)),
      t = r[0];
    return r.length !== 1 ||
      t.kind !== "call" ||
      validateLetLambdaCall(t) !== "LAMBDA"
      ? undefined
      : t.args
          .slice(0, -1)
          .map((n) =>
            validateBindingName(n, true).name.replace(/^_xlpm\./i, ""),
          );
  } catch {
    return;
  }
}
const supportedFunctionNames = [
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
  ],
  extensionRuntimeLimits = Object.freeze({
    extensions: 8,
    functions: 40,
    calls: 100,
    inputCells: 2e3,
    outputCells: 2e3,
    metadataCharacters: 2e6,
    timeoutMs: 15e3,
    maxTimeoutMs: 6e4,
  }),
  boundedTextSchema = (e) =>
    string()
      .min(1)
      .max(e)
      .refine((r) => !/[\u0000-\u001f\u007f]/.test(r)),
  safeIdentifierSchema = string()
    .regex(/^[A-Za-z][A-Za-z0-9_-]{0,63}$/)
    .refine((e) => !["constructor", "prototype", "__proto__"].includes(e)),
  extensionScalarKindSchema = _enum(["number", "string", "boolean", "scalar"]),
  extensionScalarTypeSchema = strictObject({
    type: extensionScalarKindSchema,
    nullable: boolean().optional(),
  }),
  extensionMatrixTypeSchema = strictObject({
    type: literal("matrix"),
    items: extensionScalarTypeSchema,
    maxRows: number().int().min(1).max(100),
    maxColumns: number().int().min(1).max(100),
  }),
  extensionEntityTypeSchema = strictObject({
    type: literal("entity"),
    fields: record(safeIdentifierSchema, extensionScalarTypeSchema).refine(
      (e) => Object.keys(e).length > 0 && Object.keys(e).length <= 20,
    ),
  }),
  extensionValueTypeSchema = union([
    extensionScalarTypeSchema,
    extensionMatrixTypeSchema,
    extensionEntityTypeSchema,
  ]),
  extensionPermissionSchema = _enum([
    "read-arguments",
    "return-results",
    "linked-data",
    "remote-compute",
  ]),
  extensionManifestSchema = strictObject({
    schemaVersion: literal(1),
    id: safeIdentifierSchema,
    label: boundedTextSchema(120),
    version: string()
      .regex(/^\d+\.\d+\.\d+(?:-[A-Za-z0-9.-]+)?$/)
      .max(64),
    namespace: string()
      .regex(/^[A-Z][A-Z0-9_]{1,23}$/)
      .refine(
        (e) => !["SQL", "REGISTER", "__LG", "_XLFN", "_XLUDF"].includes(e),
      ),
    runtime: _enum(["host", "office-addin", "python", "r", "linked-data"]),
    permissions: array(extensionPermissionSchema).min(2).max(4),
    recalculation: strictObject({
      mode: literal("manual"),
      invalidateOn: literal("argument-values"),
      volatile: literal(false),
    }),
    functions: array(
      strictObject({
        id: string().regex(/^[A-Z][A-Z0-9_]{0,39}$/),
        description: boundedTextSchema(300),
        parameters: array(
          strictObject({
            name: safeIdentifierSchema,
            schema: union([
              extensionScalarTypeSchema,
              extensionMatrixTypeSchema,
            ]),
          }),
        ).max(16),
        result: extensionValueTypeSchema,
      }),
    )
      .min(1)
      .max(20),
  });
function validateExtensionManifest(e) {
  const r = extensionManifestSchema.parse(e);
  if (
    new Set(r.permissions).size !== r.permissions.length ||
    !r.permissions.includes("read-arguments") ||
    !r.permissions.includes("return-results")
  )
    throw Error(
      "Extensions must declare distinct read-arguments and return-results permissions.",
    );
  if (r.runtime !== "host" && !r.permissions.includes("remote-compute"))
    throw Error("External runtimes must declare remote-compute.");
  if (new Set(r.functions.map((t) => t.id)).size !== r.functions.length)
    throw Error("Duplicate custom function names.");
  for (const t of r.functions) {
    if (supportedFunctionNames.includes(`${r.namespace}.${t.id}`))
      throw Error("An extension cannot shadow a built-in function.");
    if (new Set(t.parameters.map((n) => n.name)).size !== t.parameters.length)
      throw Error("Duplicate extension parameter names.");
    if (t.result.type === "entity" && !r.permissions.includes("linked-data"))
      throw Error("Entity results require the linked-data permission.");
  }
  return r;
}
function validateExtensionValue(e, r) {
  const t = (n, o) => {
    if (n === null && o.nullable) return null;
    if (
      (typeof n == "number" &&
        Number.isFinite(n) &&
        (o.type === "number" || o.type === "scalar")) ||
      (typeof n == "string" &&
        n.length <= 32767 &&
        (o.type === "string" || o.type === "scalar")) ||
      (typeof n == "boolean" && (o.type === "boolean" || o.type === "scalar"))
    )
      return n;
    throw Error("Extension value does not match its declared schema.");
  };
  if (r.type === "matrix") {
    if (
      !Array.isArray(e) ||
      !e.length ||
      e.length > r.maxRows ||
      !Array.isArray(e[0]) ||
      !e[0].length ||
      e[0].length > r.maxColumns ||
      e.length * e[0].length > extensionRuntimeLimits.outputCells
    )
      throw Error("Extension matrix exceeds its declared shape or cell limit.");
    const n = e[0].length;
    return e.map((o) => {
      if (!Array.isArray(o) || o.length !== n)
        throw Error("Extension matrix must be rectangular.");
      return o.map((a) => t(a, r.items));
    });
  }
  if (r.type === "entity") {
    const n = strictObject({
      kind: literal("entity"),
      id: boundedTextSchema(200),
      label: boundedTextSchema(300),
      fields: record(safeIdentifierSchema, unknown()),
      source: boundedTextSchema(1e3),
      asOf: isoDateTimeSchemaFactory(),
    }).parse(e);
    if (
      Object.keys(n.fields).sort().join("|") !==
      Object.keys(r.fields).sort().join("|")
    )
      throw Error(
        "Linked entity fields must match the declared schema exactly.",
      );
    return {
      ...n,
      fields: Object.fromEntries(
        Object.entries(r.fields).map(([o, a]) => [o, t(n.fields[o], a)]),
      ),
    };
  }
  return t(e, r);
}
const tD = new WeakSet();
function validateExtensionState(e, r) {
  if (e.extensions === undefined) return;
  const t = new Set((r.definedNames ?? []).map((u) => u.name.toUpperCase())),
    n = e.extensions;
  if (
    Array.isArray(n?.installed) &&
    n.installed.some((u) =>
      u?.manifest?.functions?.some((h) =>
        t.has(`${u.manifest.namespace}.${h.id}`),
      ),
    )
  )
    throw Error("Extension functions cannot shadow defined names.");
  if (tD.has(e.extensions)) {
    r.extensions = e.extensions;
    return;
  }
  if (
    JSON.stringify(e.extensions).length >
    extensionRuntimeLimits.metadataCharacters
  )
    throw Error("Extension metadata exceeds its limit.");
  const o = strictObject({
      schemaVersion: literal(1),
      installed: array(
        strictObject({ manifest: unknown(), enabled: boolean() }),
      ).max(extensionRuntimeLimits.extensions),
      caches: array(
        strictObject({
          key: string().regex(/^[a-f0-9]{64}$/),
          sheetId: boundedTextSchema(100),
          address: string()
            .regex(/^[A-Z]{1,3}[1-9]\d*$/)
            .refine((u) => {
              const [, l, c] = /^([A-Z]+)(\d+)$/.exec(u);
              return (
                Number(c) <= 1048576 &&
                [...l].reduce((h, f) => h * 26 + f.charCodeAt(0) - 64, 0) <=
                  16384
              );
            }),
          formula: boundedTextSchema(8192),
          functionName: boundedTextSchema(80),
          arguments: array(unknown()).max(16),
          manifestHash: string().regex(/^[a-f0-9]{64}$/),
          scopeHash: string().regex(/^[a-f0-9]{64}$/),
          result: unknown(),
          calculatedAt: isoDateTimeSchemaFactory(),
          providerRevision: boundedTextSchema(200),
        }),
      ).max(extensionRuntimeLimits.calls),
    }).parse(e.extensions),
    a = o.installed.map((u) => ({
      ...u,
      manifest: validateExtensionManifest(u.manifest),
    }));
  if (
    new Set(a.map((u) => u.manifest.id)).size !== a.length ||
    new Set(a.map((u) => u.manifest.namespace)).size !== a.length ||
    a.reduce((u, l) => u + l.manifest.functions.length, 0) >
      extensionRuntimeLimits.functions
  )
    throw Error(
      "Extension identities/namespaces must be unique within the function limit.",
    );
  const s = o.caches.map((u) => {
    const l = a.find((h) =>
        u.functionName.startsWith(h.manifest.namespace + "."),
      ),
      c = l?.manifest.functions.find(
        (h) => `${l.manifest.namespace}.${h.id}` === u.functionName,
      );
    if (!c || !l || u.arguments.length !== c.parameters.length)
      throw Error("Extension cache does not match an installed function.");
    if (
      u.key !==
      fingerprintExtensionCall(
        u.sheetId,
        u.address,
        u.formula,
        u.functionName,
        u.arguments,
        u.manifestHash,
      )
    )
      throw Error("Extension cache identity does not match its arguments.");
    return {
      ...u,
      arguments: u.arguments.map((h, f) =>
        validateExtensionValue(h, c.parameters[f].schema),
      ),
      result: validateExtensionValue(u.result, c.result),
    };
  });
  if (new Set(s.map((u) => u.key)).size !== s.length)
    throw Error("Duplicate extension cache entries.");
  const d = { schemaVersion: 1, installed: a, caches: s };
  (deepFreezeValue(d), tD.add(d), (r.extensions = d));
}
const hashExtensionManifest = (e) =>
    fingerprintCanonicalValue(validateExtensionManifest(e)),
  fingerprintExtensionCall = (e, r, t, n, o, a) =>
    fingerprintCanonicalValue([e, r, t, n, o, a]);
function listInstalledExtensionFunctions(e) {
  return (e.extensions?.installed ?? []).flatMap((t) =>
    t.manifest.functions.map((n) => ({
      entry: t,
      fn: n,
      name: `${t.manifest.namespace}.${n.id}`,
    })),
  );
}
const extensionCacheEntriesByWorkbook = new WeakMap();
function setExtensionCacheEntries(e, r) {
  extensionCacheEntriesByWorkbook.set(e, structuredClone(r));
}
function getExtensionCacheEntries(e) {
  return extensionCacheEntriesByWorkbook.get(e) ?? [];
}
function copyExtensionCacheEntries(e, r) {
  const t = extensionCacheEntriesByWorkbook.get(e);
  t && extensionCacheEntriesByWorkbook.set(r, t);
}
function hasExtensionCacheForSheet(e, r, t, n) {
  return getExtensionCacheEntries(e).some(
    (o) => o.id === r && o.manifestHash === t && o.sheetIds.includes(n),
  );
}
function resolveExtensionFunctionCall(e, r, t, n, o) {
  const a = listInstalledExtensionFunctions(e).find((m) => m.name === r);
  if (!a) return { error: "#NAME?" };
  const s = hashExtensionManifest(a.entry.manifest);
  if (
    !a.entry.enabled ||
    !hasExtensionCacheForSheet(e, a.entry.manifest.id, s, n)
  )
    return { error: "#NAME?" };
  if (t.length !== a.fn.parameters.length) return { error: "#VALUE!" };
  let d;
  try {
    d = t.map((m, v) => {
      const g = a.fn.parameters[v].schema;
      return validateExtensionValue(
        g.type === "matrix" && !Array.isArray(m) ? [[m]] : m,
        g,
      );
    });
  } catch {
    return { error: "#VALUE!" };
  }
  const u = String(e.sheets.find((m) => m.id === n)?.cells[o]?.raw ?? ""),
    l = fingerprintExtensionCall(n, o, u, r, d, s),
    c = e.extensions?.caches.find((m) => m.key === l);
  return c &&
    !getExtensionCacheEntries(e).some(
      (m) => m.id === a.entry.manifest.id && m.scopeHash === c.scopeHash,
    )
    ? { error: "#NAME?" }
    : c
      ? { value: c.result }
      : { error: "#PENDING!" };
}
const calculationBudgetDefaults = Object.freeze({
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
function readMonotonicClockMs() {
  return typeof performance < "u" ? performance.now() : Date.now();
}
function resolveCalculationBudget(e) {
  return {
    memoryBudgetBytes:
      (e?.memoryBudgetMiB ?? calculationBudgetDefaults.defaultMemoryBudgetMiB) *
      1024 *
      1024,
    workBudget: e?.workBudget ?? calculationBudgetDefaults.defaultWorkBudget,
  };
}
function estimateValueFootprintBytes(e, r) {
  let t = 0;
  const n = (o, a) => {
    if (!(t > r)) {
      if (a > 80) {
        t = r + 1;
        return;
      }
      if (typeof o == "string") t += 24 + o.length * 2;
      else if (!o || typeof o != "object") t += 8;
      else if (Array.isArray(o)) {
        t += 32 + o.length * 8;
        for (const s of o) if ((n(s, a + 1), t > r)) break;
      } else {
        t += 64;
        for (const s in o) {
          if (!Object.prototype.hasOwnProperty.call(o, s)) continue;
          t += 32 + s.length * 2;
          const d = Object.getOwnPropertyDescriptor(o, s);
          if (!("value" in d)) {
            t = r + 1;
            break;
          }
          if ((n(d.value, a + 1), t > r)) break;
        }
      }
    }
  };
  return (n(e, 0), t);
}
const rD = new WeakMap();
function estimateWorkbookMemoryBytes(e, r) {
  let t = 1024;
  const n = (o) => {
    t += estimateValueFootprintBytes(o, Math.max(0, r - t));
  };
  (n(e.definedNames),
    n(e.workbookLinks),
    n(e.extensions),
    n(e.calculation),
    n(e.datePolicy));
  for (const o of e.sheets) {
    n([
      o.id,
      o.name,
      o.tables,
      o.merges,
      o.hiddenRows,
      o.hiddenCols,
      o.outlineGroups,
    ]);
    const a = isDeepFrozenValue(o.cells),
      s = a ? rD.get(o.cells) : undefined;
    if (s !== undefined) t += s;
    else {
      const d = t;
      for (const u in o.cells)
        if (((t += 128 + 2 * u.length), n(o.cells[u].raw), t > r)) return t;
      a && rD.set(o.cells, t - d);
    }
    if (t > r) return t;
  }
  return t;
}
function createCalculationBudgetTracker(e) {
  const r = resolveCalculationBudget(e.calculation);
  let t = estimateWorkbookMemoryBytes(e, r.memoryBudgetBytes),
    n = 0,
    o,
    a,
    s = 0,
    d = 0,
    u = 0;
  const l = () => (
    !o &&
      t > r.memoryBudgetBytes &&
      ((o = "#LIMIT!"), (a = "Working-data memory budget exceeded")),
    !o &&
      n > r.workBudget &&
      ((o = "#LIMIT!"), (a = "Calculation work budget exceeded")),
    !o &&
      d + (u ? readMonotonicClockMs() - s : 0) >
        (e.calculation?.timeBudgetMs ?? 2500) &&
      ((o = "#TIMEOUT!"), (a = "Cooperative calculation time budget exceeded")),
    o
  );
  return {
    begin() {
      u++ === 0 && (s = readMonotonicClockMs());
    },
    end() {
      u > 0 && --u === 0 && (d += readMonotonicClockMs() - s);
    },
    check: l,
    work(c = 1) {
      return (
        o ||
        (n + Math.max(0, c) > r.workBudget
          ? ((o = "#LIMIT!"), (a = "Calculation work budget exceeded"), o)
          : ((n += Math.max(0, c)), l()))
      );
    },
    reserve(c) {
      return o || ((t += Math.max(0, c)), l());
    },
    value(c) {
      return (
        o ||
        ((t += estimateValueFootprintBytes(
          c,
          Math.max(0, r.memoryBudgetBytes - t),
        )),
        l())
      );
    },
    report: () => ({
      ...r,
      accountedBytes: t,
      workUnits: n,
      ...(o ? { stopped: o, reason: a } : {}),
    }),
  };
}
function isWithinNestingLimit(e) {
  const r = e.replace(/"(?:[^"]|"")*"|'(?:[^']|'')*'/g, "");
  let t = 0;
  for (const n of r) {
    if (
      ((n === "(" || n === "{") && t++,
      t > calculationBudgetDefaults.maxNesting)
    )
      return false;
    (n === ")" || n === "}") && t--;
  }
  return true;
}
const dateSystemSerialOffset = (e = "1900") => (e === "1904" ? 1462 : 0),
  dateSystemMaxSerial = (e = "1900") => 2958465 - dateSystemSerialOffset(e),
  defaultDatePolicy = Object.freeze({
    calendar: "gregorian",
    dateSystem: "1900",
    textDateOrder: "MDY",
    timeZone: "UTC",
  }),
  dateCalculationLimits = Object.freeze({
    maxSerial: 2958465,
    holidays: 1e3,
    businessDays: 366e3,
  });
function buildCalculationClock(e) {
  const r = e.datePolicy,
    t = e.calculationClock;
  return {
    ...(r
      ? {
          datePolicy: {
            calendar: r.calendar,
            dateSystem: r.dateSystem,
            textDateOrder: r.textDateOrder,
            timeZone: r.timeZone,
          },
        }
      : {}),
    ...(t
      ? {
          calculationClock: {
            calendar: t.calendar,
            dateSystem: t.dateSystem,
            textDateOrder: t.textDateOrder,
            timeZone: t.timeZone,
            calculatedAt: t.calculatedAt,
            localDateTime: t.localDateTime,
            utcOffsetMinutes: t.utcOffsetMinutes,
            nowSerial: t.nowSerial,
            todaySerial: t.todaySerial,
          },
        }
      : {}),
  };
}
function workbookUsesNowOrToday(e) {
  const r = (n) =>
    /\b(?:NOW|TODAY)\s*\(/i.test(n.replace(/"(?:[^"]|"")*"/g, '""'));
  return (
    e.sheets.some((n) =>
      Object.values(n.cells).some(
        (o) =>
          !o.literal &&
          typeof o.raw == "string" &&
          o.raw.startsWith("=") &&
          r(o.raw),
      ),
    ) || !!e.definedNames?.some((n) => r(n.refersTo))
  );
}
const millisecondsPerDay = 864e5,
  serialZeroUtcMs = Date.UTC(1899, 11, 31);
function normalizeDatePolicy(e = defaultDatePolicy) {
  const r = e;
  if (
    !r ||
    typeof r != "object" ||
    Array.isArray(r) ||
    Object.keys(r).some(
      (t) =>
        !["calendar", "dateSystem", "textDateOrder", "timeZone"].includes(t),
    ) ||
    r.calendar !== "gregorian" ||
    !["1900", "1904"].includes(r.dateSystem) ||
    !["MDY", "DMY"].includes(r.textDateOrder) ||
    typeof r.timeZone != "string" ||
    r.timeZone.length > 100 ||
    !/^[A-Za-z_]+(?:\/[A-Za-z0-9_+.-]+)*$/.test(r.timeZone)
  )
    throw Error(
      "Unsupported date policy. Use the Gregorian calendar, 1900 or 1904 serial system, MDY or DMY text dates, and a valid IANA time zone (or UTC).",
    );
  try {
    const t = new Intl.DateTimeFormat("en-US", {
      timeZone: r.timeZone,
    }).resolvedOptions().timeZone;
    return {
      calendar: "gregorian",
      dateSystem: r.dateSystem,
      textDateOrder: r.textDateOrder,
      timeZone: t,
    };
  } catch {
    throw Error(
      `Unsupported time zone: ${r.timeZone}. Use an IANA time zone or UTC.`,
    );
  }
}
const isExcelLeapYear = (e) =>
    e === 1900 || (e % 4 === 0 && (e % 100 !== 0 || e % 400 === 0)),
  daysInExcelMonth = (e, r) =>
    r === 2
      ? isExcelLeapYear(e)
        ? 29
        : 28
      : [4, 6, 9, 11].includes(r)
        ? 30
        : 31;
function dateToSerial(e, r, t, n = "1900") {
  if (n === "1904")
    return (Date.UTC(e, r - 1, t) - Date.UTC(1904, 0, 1)) / millisecondsPerDay;
  if (e === 1900 && r === 2 && t === 29) return 60;
  const o = Date.UTC(e, r - 1, t);
  return (
    (o - serialZeroUtcMs) / millisecondsPerDay +
    (o >= Date.UTC(1900, 2, 1) ? 1 : 0)
  );
}
function serialToDateParts(e, r = "1900") {
  const t = Math.floor(e);
  if (r === "1900" && t === 0) return { year: 1900, month: 1, day: 0 };
  if (r === "1900" && t === 60) return { year: 1900, month: 2, day: 29 };
  const n = new Date(
    r === "1904"
      ? Date.UTC(1904, 0, 1) + t * millisecondsPerDay
      : serialZeroUtcMs + (t - (t > 60 ? 1 : 0)) * millisecondsPerDay,
  );
  return {
    year: n.getUTCFullYear(),
    month: n.getUTCMonth() + 1,
    day: n.getUTCDate(),
  };
}
function nD(e) {
  const r = /^(\d{1,2}):(\d{2})(?::(\d{2}(?:\.\d{1,9})?))?\s*(AM|PM)?$/i.exec(
    e,
  );
  if (!r) return;
  let t = Number(r[1]);
  const n = Number(r[2]),
    o = Number(r[3] ?? 0),
    a = r[4]?.toUpperCase();
  if (!(n > 59 || o >= 60 || (a ? t < 1 || t > 12 : t > 23)))
    return (
      a && (t = (t % 12) + (a === "PM" ? 12 : 0)),
      (t * 3600 + n * 60 + o) / 86400
    );
}
function parseDateText(e, r = defaultDatePolicy) {
  const t = e.trim(),
    n =
      /^(?:(\d{4})-(\d{1,2})-(\d{1,2})|(\d{1,2})\/(\d{1,2})\/(\d{4}))(?:[ T](.+))?$/.exec(
        t,
      );
  if (!n) return nD(t);
  const o = Number(n[1] ?? n[6]),
    a = Number(n[2] ?? n[r.textDateOrder === "MDY" ? 4 : 5]),
    s = Number(n[3] ?? n[r.textDateOrder === "MDY" ? 5 : 4]);
  if (
    o < (r.dateSystem === "1904" ? 1904 : 1900) ||
    o > 9999 ||
    a < 1 ||
    a > 12 ||
    s < 1 ||
    s > daysInExcelMonth(o, a)
  )
    return;
  const d = n[7] ? nD(n[7]) : 0;
  return d === undefined ? undefined : dateToSerial(o, a, s, r.dateSystem) + d;
}
function resolveLocalDateTime(e, r = new Date()) {
  const t = normalizeDatePolicy(e),
    n = r.toISOString(),
    o = new Intl.DateTimeFormat("en-US-u-ca-gregory-nu-latn", {
      timeZone: t.timeZone,
      year: "numeric",
      month: "2-digit",
      day: "2-digit",
      hour: "2-digit",
      minute: "2-digit",
      second: "2-digit",
      hourCycle: "h23",
    }).formatToParts(r),
    a = (v) => Number(o.find((b) => b.type === v)?.value),
    s = a("year"),
    d = a("month"),
    u = a("day"),
    l = a("hour"),
    c = a("minute"),
    h = a("second"),
    f = r.getUTCMilliseconds(),
    p = dateToSerial(s, d, u, t.dateSystem);
  if (p < 0 || p > dateSystemMaxSerial(t.dateSystem))
    throw Error(
      "Calculation clock is outside the supported 1900–9999 date range.",
    );
  const m = `${s}-${String(d).padStart(2, "0")}-${String(u).padStart(2, "0")}T${String(l).padStart(2, "0")}:${String(c).padStart(2, "0")}:${String(h).padStart(2, "0")}.${String(f).padStart(3, "0")}`;
  return Object.freeze({
    ...t,
    calculatedAt: n,
    localDateTime: m,
    utcOffsetMinutes: (Date.UTC(s, d - 1, u, l, c, h, f) - r.getTime()) / 6e4,
    todaySerial: p,
    nowSerial: p + (l * 3600 + c * 60 + h + f / 1e3) / 86400,
  });
}
const calculationFeatureSupport = Object.freeze({
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
function findVolatileFunctionNames(e) {
  const r = e.replace(/"(?:[^"]|"")*"|'(?:[^']|'')*'(?=\s*!)/g, "");
  return [
    ...new Set(
      [...r.matchAll(/\b(?:_xlfn\.)?(RAND|RANDBETWEEN|NOW|TODAY)\s*\(/gi)].map(
        (t) => t[1].toUpperCase(),
      ),
    ),
  ];
}
function workbookHasVolatileFunctions(e) {
  return (
    !!e.definedNames?.some(
      (t) => findVolatileFunctionNames(t.refersTo).length,
    ) ||
    e.sheets.some((t) =>
      Object.values(t.cells).some(
        (n) =>
          !n.literal &&
          typeof n.raw == "string" &&
          n.raw.startsWith("=") &&
          findVolatileFunctionNames(n.raw).length,
      ),
    )
  );
}
const dateFunctionNames = [
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
  notDateFunctionSentinel = Symbol("not-date-function"),
  UJ = {
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
function evaluateDateFunction(e, r, t) {
  if (!dateFunctionNames.includes(e)) return notDateFunctionSentinel;
  const n = t.clock.dateSystem,
    o = dateSystemMaxSerial(n),
    a = (w, E, A) => dateToSerial(w, E, A, n),
    s = (w) => serialToDateParts(w, n),
    d = (w) => (((w + dateSystemSerialOffset(n) + 5) % 7) + 7) % 7,
    u = (w = "#VALUE!") => {
      throw t.issue(w);
    },
    [l, c] = UJ[e];
  (r.length < l || r.length > c) && u();
  const h = r.map(t.retrieve),
    f = (w) => {
      if (t.isIssue(w)) throw w;
      return Array.isArray(w) ? (w.length === 1 ? f(w[0]) : u()) : w;
    },
    p = (w) => {
      if (((w = f(w)), w == null)) return 0;
      if (typeof w == "boolean") return Number(w);
      typeof w != "number" &&
        (typeof w != "string" ||
          !/^\s*[+-]?(?:\d+\.?\d*|\.\d+)(?:E[+-]?\d+)?\s*$/i.test(w)) &&
        u();
      const E = Number(w);
      return Number.isFinite(E) ? E : u("#NUM!");
    },
    m = (w) => (Number.isFinite(w) && w >= 0 && w < o + 1 ? w : u("#NUM!")),
    v = (w) => {
      w = f(w);
      const E =
        typeof w == "string" &&
        !/^\s*[+-]?(?:\d+\.?\d*|\.\d+)(?:E[+-]?\d+)?\s*$/i.test(w)
          ? parseDateText(w, t.clock)
          : p(w);
      return E === undefined ? u() : m(E);
    },
    g = (w) => Math.floor(v(w)),
    b = (w, E, A = false) => {
      const k = s(w),
        F = k.year * 12 + k.month - 1 + E,
        I = Math.floor(F / 12),
        O = (((F % 12) + 12) % 12) + 1;
      return (
        (I < (n === "1904" ? 1904 : 1900) || I > 9999) && u("#NUM!"),
        m(
          a(
            I,
            O,
            A
              ? daysInExcelMonth(I, O)
              : Math.min(k.day, daysInExcelMonth(I, O)),
          ),
        )
      );
    };
  if (e === "TODAY") return t.clock.todaySerial;
  if (e === "NOW") return t.clock.nowSerial;
  if (e === "DATE") {
    let w = Math.trunc(p(h[0]));
    const E = Math.trunc(p(h[1])),
      A = Math.trunc(p(h[2]));
    ((w < 0 || w >= 1e4) && u("#NUM!"), w < 1900 && (w += 1900));
    const k = w * 12 + E - 1;
    w = Math.floor(k / 12);
    const F = (((k % 12) + 12) % 12) + 1;
    return (
      (w < 1899 || w > 1e4 || Math.abs(A) > dateCalculationLimits.maxSerial) &&
        u("#NUM!"),
      m(a(w, F, 1) + A - 1)
    );
  }
  if (e === "DATEVALUE" || e === "TIMEVALUE") {
    const w = f(h[0]);
    typeof w != "string" && u();
    const E = parseDateText(w, t.clock);
    return (
      E === undefined && u(),
      e === "DATEVALUE" &&
        !/^(?:\d{4}-\d{1,2}-\d{1,2}|\d{1,2}\/\d{1,2}\/\d{4})(?:[ T]|$)/.test(
          w.trim(),
        ) &&
        u(),
      e === "DATEVALUE" ? Math.floor(E) : E - Math.floor(E)
    );
  }
  if (e === "TIME") {
    const [w, E, A] = h.map((k) => Math.trunc(p(k)));
    return (
      [w, E, A].some((k) => k < 0 || k > 32767) && u("#NUM!"),
      ((w * 3600 + E * 60 + A) % 86400) / 86400
    );
  }
  if (["HOUR", "MINUTE", "SECOND"].includes(e)) {
    const w = v(h[0]),
      E = Math.max(1e-9, Number.EPSILON * w * 86400),
      A = Math.floor((w - Math.floor(w)) * 86400 + E) % 86400;
    return e === "HOUR"
      ? Math.floor(A / 3600)
      : e === "MINUTE"
        ? Math.floor(A / 60) % 60
        : A % 60;
  }
  if (["DAY", "MONTH", "YEAR"].includes(e)) return s(g(h[0]))[e.toLowerCase()];
  if (e === "EDATE" || e === "EOMONTH")
    return b(g(h[0]), Math.trunc(p(h[1])), e === "EOMONTH");
  if (e === "DAYS") return g(h[0]) - g(h[1]);
  if (e === "WEEKDAY") {
    const w = g(h[0]),
      E = h[1] == null ? 1 : Math.trunc(p(h[1])),
      A = d(w);
    if (E === 3) return A;
    const k =
      E === 1 || E === 17
        ? 6
        : E === 2 || E === 11
          ? 0
          : E >= 12 && E <= 16
            ? E - 11
            : u("#NUM!");
    return ((A - k + 7) % 7) + 1;
  }
  if (e === "WEEKNUM" || e === "ISOWEEKNUM") {
    const w = g(h[0]),
      E = e === "ISOWEEKNUM" ? 21 : h[1] == null ? 1 : Math.trunc(p(h[1]));
    if (E === 21) {
      const F = w + 3 - d(w),
        I = s(F).year;
      return Math.floor((F - a(I, 1, 1)) / 7) + 1;
    }
    const A =
        E === 1 || E === 17
          ? 6
          : E === 2 || E === 11
            ? 0
            : E >= 12 && E <= 16
              ? E - 11
              : u("#NUM!"),
      k = a(s(w).year, 1, 1);
    return Math.floor((w - k + ((d(k) - A + 7) % 7)) / 7) + 1;
  }
  if (e === "DATEDIF") {
    const w = g(h[0]),
      E = g(h[1]),
      A = String(f(h[2])).toUpperCase();
    w > E && u("#NUM!");
    const k = s(w),
      F = s(E),
      I = (F.year - k.year) * 12 + F.month - k.month - (F.day < k.day ? 1 : 0);
    if (A === "D") return E - w;
    if (A === "Y") return Math.floor(I / 12);
    if (A === "M") return I;
    if (A === "YM") return I % 12;
    if (A === "MD") return E - b(w, I);
    if (A === "YD") {
      const O =
        F.year -
        (F.month < k.month || (F.month === k.month && F.day < k.day) ? 1 : 0);
      return E - a(O, k.month, Math.min(k.day, daysInExcelMonth(O, k.month)));
    }
    u("#NUM!");
  }
  const y = (w, E, A) => {
    const k = s(w),
      F = s(E);
    let I = k.day,
      O = F.day;
    return (
      A
        ? ((I = Math.min(I, 30)), (O = Math.min(O, 30)))
        : (I === daysInExcelMonth(k.year, k.month) && (I = 30),
          O === 31 && I >= 30 && (O = 30),
          k.month === 2 &&
            k.day === daysInExcelMonth(k.year, 2) &&
            F.month === 2 &&
            O === daysInExcelMonth(F.year, 2) &&
            (O = 30)),
      (F.year - k.year) * 360 + (F.month - k.month) * 30 + O - I
    );
  };
  if (e === "DAYS360") {
    const w = f(h[2] ?? false);
    return (
      typeof w != "boolean" && typeof w != "number" && u(),
      y(g(h[0]), g(h[1]), !!w)
    );
  }
  if (e === "YEARFRAC") {
    let w = g(h[0]),
      E = g(h[1]);
    const A = h[2] == null ? 0 : Math.trunc(p(h[2]));
    if (((A < 0 || A > 4) && u("#NUM!"), w > E && ([w, E] = [E, w]), w === E))
      return 0;
    if (A === 0 || A === 4) return y(w, E, A === 4) / 360;
    if (A === 2 || A === 3) return (E - w) / (A === 2 ? 360 : 365);
    const k = s(w),
      F = s(E);
    let I = 365;
    const O = a(
      k.year + 1,
      k.month,
      Math.min(k.day, daysInExcelMonth(k.year + 1, k.month)),
    );
    if (k.year === F.year) I = isExcelLeapYear(k.year) ? 366 : 365;
    else if (E <= O)
      for (let j = k.year; j <= F.year; j++)
        isExcelLeapYear(j) && a(j, 2, 29) >= w && a(j, 2, 29) <= E && (I = 366);
    else I = (a(F.year + 1, 1, 1) - a(k.year, 1, 1)) / (F.year - k.year + 1);
    return (E - w) / I;
  }
  if (e.startsWith("WORKDAY") || e.startsWith("NETWORKDAYS")) {
    const w = e.startsWith("NETWORK"),
      E = e.endsWith(".INTL"),
      A = g(h[0]),
      k = f(E ? (h[2] ?? 1) : 1);
    let F;
    if (typeof k == "string") (/^[01]{7}$/.test(k) || u(), (F = k));
    else {
      const $ = p(k);
      let U;
      if ($ >= 1 && $ <= 7 && Number.isInteger($))
        U = [($ + 4) % 7, ($ + 5) % 7];
      else if ($ >= 11 && $ <= 17 && Number.isInteger($)) U = [($ - 5) % 7];
      else return u("#NUM!");
      F = Array.from({ length: 7 }, (G, H) => (U.includes(H) ? "1" : "0")).join(
        "",
      );
    }
    const I = ($) => (Array.isArray($) ? $.flatMap(I) : [$]),
      O = h[E ? 3 : 2] === undefined ? [] : I(h[E ? 3 : 2]);
    O.length > dateCalculationLimits.holidays && u("#LIMIT!");
    const j = new Set();
    for (const $ of O)
      $ == null || $ === "" || (typeof $ == "boolean" && u(), j.add(g($)));
    const z = [...F].filter(($) => $ === "0").length,
      V = ($) => F[d($)] === "0",
      D = ($, U) => {
        if ($ > U) return 0;
        const G = U - $ + 1;
        let H = Math.floor(G / 7) * z;
        for (let Q = 0; Q < G % 7; Q++) V($ + Q) && H++;
        for (const Q of j) Q >= $ && Q <= U && V(Q) && H--;
        return H;
      };
    if (w) {
      let $ = g(h[1]);
      return (
        Math.abs($ - A) > dateCalculationLimits.businessDays && u("#LIMIT!"),
        $ < A ? -D($, A) : D(A, $)
      );
    }
    const B = Math.trunc(p(h[1]));
    if (
      (z || u(),
      Math.abs(B) > dateCalculationLimits.businessDays && u("#LIMIT!"),
      B === 0)
    )
      return A;
    const _ = B > 0,
      C = Math.abs(B),
      T = _ ? o - A : A,
      N = ($) => (_ ? D(A + 1, A + $) : D(A - $, A - 1));
    N(T) < C && u("#NUM!");
    let R = 1,
      L = T;
    for (; R < L;) {
      const $ = Math.floor((R + L) / 2);
      N($) >= C ? (L = $) : (R = $ + 1);
    }
    return A + (_ ? R : -R);
  }
  return u();
}
const financialSolverLimits = Object.freeze({
  cashFlows: 1e4,
  periods: 1e4,
  couponPeriods: 1e3,
  evaluations: 2048,
  minRate: -0.999999999,
  maxRate: 1e6,
  relativeResidual: 1e-11,
});

export {
  fingerprintExtensionCall,
  parseFormulaTree,
  validateLetLambdaCall,
  parseTransformationRecipe,
  dateCalculationLimits,
  copyExtensionCacheEntries,
  validateSavedRecipeMetadata,
  walkFormulaNodes,
  Gq,
  calculationBudgetDefaults,
  resolveCalculationBudget,
  isoDateTimeSchemaFactory,
  resolveExtensionFunctionCall,
  deepFreezeValue,
  extensionRuntimeLimits,
  validateBindingName,
  isWithinNestingLimit,
  normalizeFunctionName,
  buildCalculationClock,
  tokenizeFormulaText,
  estimateWorkbookMemoryBytes,
  findVolatileFunctionNames,
  validateExtensionManifest,
  createCalculationBudgetTracker,
  defaultDatePolicy,
  dropBlankTextTokens,
  dateToSerial,
  dateFunctionNames,
  evaluateDateFunction,
  setExtensionCacheEntries,
  transformationRecipeLimits,
  supportedFunctionNames,
  unwrapDefaultExport,
  hashExtensionManifest,
  workbookUsesNowOrToday,
  isDeepFrozenValue,
  getExtensionCacheEntries,
  fingerprintCanonicalValue,
  notDateFunctionSentinel,
  sha256HexDigest,
  validateExtensionState,
  estimateValueFootprintBytes,
  measureSerializedBytes,
  normalizeBindingName,
  resolveLocalDateTime,
  validateExtensionValue,
  canonicalJsonStringify,
  workbookHasVolatileFunctions,
  daysInExcelMonth,
  financialSolverLimits,
  readMonotonicClockMs,
  recipeStoreFormatId,
  hasLinkedOrBoundCells,
  serialToDateParts,
  normalizeDatePolicy,
  uEe,
  dateSystemSerialOffset,
  dateSystemMaxSerial,
  calculationFeatureSupport,
  defineDataPropertyByKey,
  listInstalledExtensionFunctions,
  parseDateText,
};
export { boundedTextSchema };
