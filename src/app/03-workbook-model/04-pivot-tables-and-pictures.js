import {
  isNonArrayObject,
  isPlainObjectValue,
  throwWorkbookError,
} from "./01-cell-styles-and-themes.js";
import { createUuid, isNullOrEmptyValue } from "./03-defined-names.js";
import { boundedTextSchema } from "../01-core-schemas-and-hashing.js";
import { flattenNestedValues } from "../02-financial-functions.js";
import {
  workbookLimits,
  supportedFinancialFunctionNames,
} from "../02-financial-functions.js";
import { strToU8 } from "fflate";
import {
  string,
  number,
  boolean,
  null as _null,
  array,
  strictObject,
  discriminatedUnion,
  literal,
  record,
} from "zod";
import {
  isoDateTimeSchemaFactory,
  deepFreezeValue,
  sha256HexDigest,
  supportedFunctionNames,
  normalizeFunctionName,
  parseFormulaTree,
} from "../01-core-schemas-and-hashing.js";
import { assertSheetUnprotected } from "./01-cell-styles-and-themes.js";
import {
  FormulaParser,
  statisticalFunctionNames,
  XL,
  specialistFunctionNames,
  significanceTestFunctionNames,
  ownedLookupFunctionNames,
  modernFunctionNames,
} from "./02-number-formats.js";
import {
  definedNameScopeKey,
  hashPictureBytes,
  maxPictureHashBytes,
  stripFormulaStringLiterals,
  replaceTextSpans,
  localBindingFunctionNames,
  m5,
  validateNameFormulaText,
  formatQuotedSheetPrefix,
  validateDrawingText,
} from "./03-defined-names.js";

function inspectPictureDataUrl(e) {
  var s;
  if (
    typeof e != "string" ||
    e.length > Math.ceil(maxPictureHashBytes / 3) * 4 + 30
  )
    throw Error("Pictures must be local PNG/JPEG data of at most 2 MB.");
  const r = /^data:image\/(png|jpeg);base64,([A-Za-z0-9+/]+={0,2})$/.exec(e);
  if (!r || r[2].length % 4 !== 0)
    throw Error("Pictures require canonical local PNG/JPEG base64 data.");
  let t;
  try {
    const d = atob(r[2]);
    if (btoa(d) !== r[2]) throw Error();
    t = Uint8Array.from(d, (u) => u.charCodeAt(0));
  } catch {
    throw Error("The picture data is invalid.");
  }
  if (t.length > maxPictureHashBytes)
    throw Error("Pictures must be at most 2 MB.");
  let n = 0,
    o = 0;
  const a = () => {
    if (!n || !o || n > 1e4 || o > 1e4 || n * o > 4e7)
      throw Error(
        "Picture dimensions are invalid or exceed 40 megapixels / 10,000 pixels per side.",
      );
  };
  if (r[1] === "png") {
    if (
      t.length < 45 ||
      ![137, 80, 78, 71, 13, 10, 26, 10].every((p, m) => t[m] === p)
    )
      throw Error("The file contents do not match a PNG picture.");
    const d = new DataView(t.buffer, t.byteOffset, t.byteLength);
    if (
      d.getUint32(8) !== 13 ||
      String.fromCharCode(...t.subarray(12, 16)) !== "IHDR"
    )
      throw Error("PNG requires an intact first IHDR chunk.");
    if (
      ((n = d.getUint32(16)),
      (o = d.getUint32(20)),
      a(),
      !(
        (s = {
          0: [1, 2, 4, 8, 16],
          2: [8, 16],
          3: [1, 2, 4, 8],
          4: [8, 16],
          6: [8, 16],
        }[t[25]]) != null && s.includes(t[24])
      ) ||
        t[26] !== 0 ||
        t[27] !== 0 ||
        t[28] > 1)
    )
      throw Error("Unsupported PNG pixel encoding.");
    let l = 8,
      c = 0,
      h = false,
      f = false;
    for (; l < t.length;) {
      if (++c > 4096 || l + 12 > t.length)
        throw Error("PNG chunk structure is invalid or excessive.");
      const p = d.getUint32(l),
        m = String.fromCharCode(...t.subarray(l + 4, l + 8));
      if (p > t.length - l - 12 || !/^[A-Za-z]{4}$/.test(m))
        throw Error("PNG chunk length or type is invalid.");
      if (m === "IHDR" && l !== 8) throw Error("PNG has duplicate dimensions.");
      if ((m === "IDAT" && p > 0 && (h = true), m === "IEND")) {
        if (p !== 0 || l + 12 !== t.length)
          throw Error("PNG has invalid trailing data.");
        f = true;
      }
      l += p + 12;
    }
    if (!h || !f) throw Error("PNG data or end marker is missing.");
  } else {
    if (t.length < 12 || t[0] !== 255 || t[1] !== 216)
      throw Error("The file contents do not match a JPEG picture.");
    let d = 2,
      u = 0,
      l = false,
      c = false;
    for (; d < t.length;) {
      if (++u > 65536) throw Error("JPEG has too many segments.");
      if (l) for (; d < t.length && t[d] !== 255;) d++;
      if (d >= t.length || t[d++] !== 255)
        throw Error("JPEG marker structure is invalid.");
      for (; t[d] === 255;) d++;
      const h = t[d++];
      if (l && (h === 0 || (h >= 208 && h <= 215))) continue;
      if (h === 217) {
        if (d !== t.length) throw Error("JPEG has trailing data.");
        c = true;
        break;
      }
      if (h === 1) continue;
      if (d + 2 > t.length) throw Error("JPEG segment is truncated.");
      const f = t[d] * 256 + t[d + 1];
      if (f < 2 || d + f > t.length)
        throw Error("JPEG segment length is invalid.");
      if (h >= 192 && h <= 207 && ![196, 200, 204].includes(h)) {
        if (![192, 193, 194].includes(h) || f < 8 || t[d + 2] !== 8)
          throw Error("Unsupported JPEG pixel encoding.");
        const p = t[d + 7];
        if (p < 1 || p > 4 || f !== 8 + p * 3)
          throw Error("JPEG frame components are invalid.");
        const m = t[d + 3] * 256 + t[d + 4],
          v = t[d + 5] * 256 + t[d + 6];
        if (n && (n !== v || o !== m)) throw Error("JPEG dimensions conflict.");
        ((n = v), (o = m), a());
      }
      if (h === 218) {
        if (!n || f < 6 || f !== 6 + t[d + 2] * 2)
          throw Error("JPEG scan is invalid.");
        l = true;
      }
      d += f;
    }
    if (!l || !c) throw Error("JPEG scan or end marker is missing.");
    a();
  }
  return { bytes: t, extension: r[1], width: n, height: o };
}
function assertPlainTextLabel(e, r, t) {
  const n = validateDrawingText(e, r, t);
  if (!n.trim() || /[\u0000-\u001f\u007f]/.test(n))
    throw Error(
      `Invalid picture ${t}: use nonblank plain text without control characters.`,
    );
  return n;
}
function validatePictureEvidence(e, r) {
  if (
    !isNonArrayObject(e) ||
    ![Object.prototype, null].includes(Object.getPrototypeOf(e)) ||
    Reflect.ownKeys(e).some(
      (o) =>
        typeof o != "string" ||
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
        ].includes(o),
    )
  )
    throw Error("Invalid picture evidence metadata.");
  if (
    e.version !== 1 ||
    typeof e.origin != "string" ||
    !["local-file", "xlsx-media", "legacy-native"].includes(e.origin)
  )
    throw Error("Invalid picture evidence origin.");
  const t = hashPictureBytes(r);
  if (
    typeof e.sha256 != "string" ||
    !/^[a-f0-9]{64}$/.test(e.sha256) ||
    e.sha256 !== t ||
    e.sourceReference !== `sha256:${t}` ||
    e.byteLength !== r.length
  )
    throw Error(
      "Picture evidence digest or byte length does not match the embedded image.",
    );
  const n = {
    version: 1,
    origin: e.origin,
    sourceReference: `sha256:${t}`,
    sha256: t,
    byteLength: r.length,
  };
  for (const o of [
    "filename",
    "sourceLocator",
    "containerFilename",
    "mediaReference",
  ])
    e[o] !== undefined &&
      (n[o] = assertPlainTextLabel(e[o], o === "sourceLocator" ? 2e3 : 512, o));
  if (n.origin === "local-file" && !n.filename)
    throw Error("A local evidence picture requires its original filename.");
  if (n.origin !== "xlsx-media" && (n.containerFilename || n.mediaReference))
    throw Error("Only observed XLSX media may have container metadata.");
  if (n.origin === "legacy-native" && n.filename)
    throw Error(
      "An original filename cannot be inferred for a legacy picture.",
    );
  return n;
}
function createPictureEvidence(e, r) {
  const { bytes } = inspectPictureDataUrl(e),
    n = hashPictureBytes(bytes);
  return validatePictureEvidence(
    {
      ...r,
      version: 1,
      sourceReference: `sha256:${n}`,
      sha256: n,
      byteLength: bytes.length,
    },
    bytes,
  );
}
function F2(e, r, t) {
  if (!e.sheets.some((n) => n.id === r))
    throw Error("The worksheet no longer exists.");
  return {
    workbookId: e.id,
    revision: e.revision,
    sheetId: r,
    ...(t !== undefined ? { selection: t } : {}),
    guard: JSON.stringify(e),
  };
}
function assertDrawingPreviewCurrent(e, r, t = {}) {
  if (t.readOnly) throw Error("This workbook is read-only.");
  const n = e.sheets.find((a) => a.id === r);
  if (!n) throw Error("The worksheet no longer exists.");
  assertSheetUnprotected(n);
  const o = t.snapshot;
  if (
    o &&
    (o.workbookId !== e.id ||
      o.revision !== e.revision ||
      o.sheetId !== r ||
      o.guard !== JSON.stringify(e) ||
      o.selection !== t.selection)
  )
    throw Error(
      "The drawing preview is stale. Reopen it or select the object again.",
    );
}
const throwErrorMessage = (e) => {
    throw Error(e);
  },
  isFiniteNumber = (e) => typeof e == "number" && Number.isFinite(e);
function validatePivotRefreshEvidence(e) {
  isNonArrayObject(e) || throwErrorMessage("Invalid pivot refresh evidence.");
  const r = e;
  (typeof r.at != "string" ||
    !/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}Z$/.test(r.at) ||
    !Number.isFinite(Date.parse(r.at)) ||
    new Date(r.at).toISOString() !== r.at ||
    typeof r.sourceSheetName != "string" ||
    !r.sourceSheetName ||
    r.sourceSheetName.length > 31 ||
    typeof r.sourceRange != "string" ||
    !/^[A-Z]+\d+:[A-Z]+\d+$/.test(r.sourceRange)) &&
    throwErrorMessage("Invalid pivot refresh time or source label.");
  for (const n of ["sourceFingerprint", "outputFingerprint"])
    (typeof r[n] != "string" || !/^[a-f0-9]{64}$/.test(r[n])) &&
      throwErrorMessage("Invalid pivot content fingerprint.");
  for (const n of ["sourceRows", "includedRows", "excludedRows"])
    (!Number.isSafeInteger(r[n]) || Number(r[n]) < 0 || Number(r[n]) > 1e5) &&
      throwErrorMessage("Invalid pivot population counts.");
  (Number(r.sourceRows) !== Number(r.includedRows) + Number(r.excludedRows) ||
    !Array.isArray(r.measures) ||
    !r.measures.length ||
    r.measures.length > 100) &&
    throwErrorMessage("Pivot population evidence does not reconcile.");
  const t = r.measures.map((n) => {
    (!isNonArrayObject(n) ||
      typeof n.name != "string" ||
      !n.name ||
      n.name.length > 512 ||
      typeof n.field != "string" ||
      !n.field ||
      n.field.length > 255 ||
      !["sum", "count", "average", "min", "max", "distinctCount"].includes(
        String(n.operation),
      ) ||
      !(
        n.result === null ||
        typeof n.result == "boolean" ||
        isFiniteNumber(n.result) ||
        (typeof n.result == "string" && n.result.length <= 32767)
      )) &&
      throwErrorMessage("Invalid pivot measure evidence.");
    const o = n;
    for (const a of ["numericCount", "nonblankCount"])
      (!Number.isSafeInteger(o[a]) ||
        Number(o[a]) < 0 ||
        Number(o[a]) > Number(r.includedRows)) &&
        throwErrorMessage("Invalid pivot measure count.");
    return (
      Number(o.numericCount) > Number(o.nonblankCount) &&
        throwErrorMessage("Invalid pivot numeric count."),
      {
        name: o.name,
        field: o.field,
        operation: o.operation,
        result: o.result,
        numericCount: o.numericCount,
        nonblankCount: o.nonblankCount,
      }
    );
  });
  return {
    at: r.at,
    sourceSheetName: r.sourceSheetName,
    sourceRange: r.sourceRange,
    sourceFingerprint: r.sourceFingerprint,
    outputFingerprint: r.outputFingerprint,
    sourceRows: r.sourceRows,
    includedRows: r.includedRows,
    excludedRows: r.excludedRows,
    measures: t,
  };
}
function bucketPivotDateValue(e, r, t) {
  if (isNullOrEmptyValue(e)) return null;
  let n;
  if (r.dateSystem === "iso") {
    if (typeof e != "string" || !/^\d{4}-\d{2}-\d{2}$/.test(e))
      return throwErrorMessage(
        `Date grouping for “${r.field}” needs YYYY-MM-DD text on source row ${t + 1}; choose the correct date system or clean the source.`,
      );
    if (
      ((n = new Date(e + "T00:00:00.000Z")),
      !Number.isFinite(n.getTime()) ||
        n.toISOString().slice(0, 10) !== e ||
        e.startsWith("0000"))
    )
      return throwErrorMessage(
        `Invalid date in “${r.field}” on source row ${t + 1}.`,
      );
  } else {
    if (
      !isFiniteNumber(e) ||
      e < (r.dateSystem === "1900" ? 1 : 0) ||
      e > 2958465
    )
      return throwErrorMessage(
        `Date grouping for “${r.field}” requires a valid ${r.dateSystem} numeric date serial on source row ${t + 1}.`,
      );
    const s = Math.floor(e),
      d =
        r.dateSystem === "1900" ? Date.UTC(1899, 11, 31) : Date.UTC(1904, 0, 1);
    if (
      ((n = new Date(
        d + (s - (r.dateSystem === "1900" && s >= 60 ? 1 : 0)) * 864e5,
      )),
      n.getUTCFullYear() > 9999)
    )
      return throwErrorMessage("Grouped date exceeds year 9999.");
  }
  const o = String(n.getUTCFullYear()).padStart(4, "0"),
    a = n.getUTCMonth() + 1;
  return r.unit === "year"
    ? o
    : r.unit === "quarter"
      ? `${o} Q${Math.ceil(a / 3)}`
      : `${o}-${String(a).padStart(2, "0")}`;
}
const createAggregateAccumulator = (e = false) => ({
  count: 0,
  numeric: 0,
  sum: 0,
  min: 1 / 0,
  max: -1 / 0,
  ...(e ? { distinct: new Set() } : {}),
});
function SO(e, r) {
  var t;
  (isNullOrEmptyValue(r) ||
    (e.count++,
    (t = e.distinct) == null ||
      t.add(
        JSON.stringify([
          typeof r,
          typeof r == "string" ? r.toLocaleLowerCase("en-US") : r,
        ]),
      )),
    isFiniteNumber(r) &&
      (e.numeric++,
      (e.sum += r),
      (e.min = Math.min(e.min, r)),
      (e.max = Math.max(e.max, r)),
      Number.isFinite(e.sum) ||
        throwErrorMessage("Aggregation exceeds finite numeric bounds.")));
}
function finalizeAggregateState(e, r) {
  return r === "count"
    ? e.count
    : r === "distinctCount"
      ? (e.distinct?.size ?? 0)
      : r === "sum"
        ? e.sum
        : e.numeric
          ? r === "average"
            ? e.sum / e.numeric
            : r === "min"
              ? e.min
              : e.max
          : null;
}
function computePivotTable(e, r, t) {
  const n = (C) => {
      const T = r.indexOf(C);
      return T < 0
        ? throwErrorMessage(
            `Source field “${C}” was not found. Refresh the field selection.`,
          )
        : T;
    },
    o = e.rowFields.map(n),
    a = e.columnFields.map(n);
  for (const C of e.dateGroups ?? [])
    [...e.rowFields, ...e.columnFields].includes(C.field) ||
      throwErrorMessage(
        `Date group “${C.field}” must be selected as a row or column field.`,
      );
  for (const C of e.calculatedFields ?? [])
    (r.some((T) => T.toLowerCase() === C.name.toLowerCase()) &&
      throwErrorMessage(
        `Calculated field “${C.name}” duplicates a source header.`,
      ),
      n(C.field),
      C.otherField && n(C.otherField));
  const s = e.values.map((C) => {
      const T = e.calculatedFields?.find((R) => R.name === C.field);
      return (
        T &&
          C.operation !== "sum" &&
          throwErrorMessage(
            `Calculated field “${C.field}” requires sum aggregation; its arithmetic is applied to source-field sums.`,
          ),
        {
          ...C,
          index: n(T?.field ?? C.field),
          calculated: T,
          other: T != null && T.otherField ? n(T.otherField) : undefined,
        }
      );
    }),
    d = () =>
      s.map((C) => ({
        ...createAggregateAccumulator(C.operation === "distinctCount"),
        ...(C.other !== undefined
          ? { right: createAggregateAccumulator() }
          : {}),
      }));
  t.length *
    s.length *
    Math.max(1, e.subtotals ? o.length : 1) *
    Math.max(1, e.subtotals ? a.length : 1) >
    2e6 &&
    throwErrorMessage(
      "Pivot aggregation exceeds the 2,000,000 contribution budget. Reduce value fields, hierarchy depth or source rows.",
    );
  let u = 0;
  const l = (C, T) =>
      s.forEach((N, R) => {
        const L = C[R].distinct?.size ?? 0;
        (SO(C[R], T[N.index]),
          (u += (C[R].distinct?.size ?? 0) - L),
          u > 25e4 &&
            throwErrorMessage(
              "Pivot distinct counts exceed 250,000 retained group entries. Reduce groups, value fields or source rows.",
            ),
          N.other !== undefined && SO(C[R].right, T[N.other]));
      }),
    c = (C, T) => {
      const N = s[T],
        R = C?.[T];
      if (!N.calculated)
        return finalizeAggregateState(
          R ?? createAggregateAccumulator(N.operation === "distinctCount"),
          N.operation,
        );
      if (!R) return null;
      const L = R.sum,
        $ = N.other !== undefined ? R.right.sum : N.calculated.constant;
      N.calculated.operation === "divide" &&
        $ === 0 &&
        throwErrorMessage(
          `Calculated field “${N.field}” divides by zero in a group or total. Filter, clean or change the calculation; no result was written.`,
        );
      const U =
        N.calculated.operation === "add"
          ? L + $
          : N.calculated.operation === "subtract"
            ? L - $
            : N.calculated.operation === "multiply"
              ? L * $
              : L / $;
      return Number.isFinite(U)
        ? U
        : throwErrorMessage(
            `Calculated field “${N.field}” exceeds finite numeric bounds.`,
          );
    },
    h = (C, T) =>
      C.map((N) => {
        const R = e.dateGroups?.find((U) => U.field === r[N]),
          L = T.values[N];
        return R
          ? bucketPivotDateValue(L, R, T.index)
          : isNullOrEmptyValue(L)
            ? null
            : L;
      }),
    f = (C) =>
      [
        C,
        ...(e.subtotals
          ? Array.from({ length: Math.max(0, C.length - 1) }, (T, N) =>
              C.slice(0, N + 1),
            )
          : []),
      ].map((T) => ({
        key: JSON.stringify(T),
        parts: T,
        subtotal: T.length < C.length,
      })),
    p = new Map(),
    m = new Map(),
    v = new Map(),
    g = new Map(),
    b = new Map(),
    y = d(),
    w = (C, T, N) => {
      (C.has(T) || C.set(T, d()), l(C.get(T), N));
    };
  for (const C of t) {
    const T = f(h(o, C)),
      N = f(h(a, C));
    (T.forEach((L) => p.set(L.key, L)), N.forEach((L) => m.set(L.key, L)));
    const R =
      Math.max(1, o.length) +
      (m.size + +(!!e.grandTotals && !!a.length)) * s.length;
    (p.size + 2) * R > 1e5 &&
      throwErrorMessage(
        "The pivot grouping would exceed the 100,000-cell output budget. Reduce group fields or filter the source.",
      );
    for (const L of T) {
      w(g, L.key, C.values);
      for (const $ of N) w(v, JSON.stringify([L.key, $.key]), C.values);
    }
    for (const L of N) w(b, L.key, C.values);
    l(y, C.values);
  }
  (m.size || m.set("[]", { key: "[]", parts: [], subtotal: false }),
    !p.size &&
      !o.length &&
      p.set("[]", { key: "[]", parts: [], subtotal: false }));
  function E(C, T) {
    var $;
    if (!e.subtotals && !(($ = e.sorts) != null && $.length))
      return [...C.values()];
    if (!T.length) return [...C.values()];
    const N = new Map();
    for (const U of C.values())
      if (!U.subtotal)
        for (let G = 0; G < U.parts.length; G++) {
          const H = JSON.stringify(U.parts.slice(0, G)),
            Q = U.parts.slice(0, G + 1),
            X = JSON.stringify(Q);
          (N.has(H) || N.set(H, new Map()), N.get(H).set(X, Q));
        }
    const R = [],
      L = (U) => {
        const G = e.sorts?.find((q) => q.field === T[U.length]),
          H = [...(N.get(JSON.stringify(U))?.values() ?? [])];
        G &&
          H.sort((q, W) => {
            const K = q.at(-1),
              te = W.at(-1),
              ae =
                K === te
                  ? 0
                  : K === null
                    ? -1
                    : te === null
                      ? 1
                      : typeof K == "number" && typeof te == "number"
                        ? K - te
                        : String(K).localeCompare(String(te), "en-US", {
                            numeric: true,
                          });
            return G.direction === "ascending" ? ae : -ae;
          });
        for (const q of H) L(q);
        const Q = C.get(JSON.stringify(U));
        Q && R.push(Q);
      };
    return (L([]), R);
  }
  const A = E(m, e.columnFields),
    k = E(p, e.rowFields),
    F = (C) =>
      C.name ??
      (C.calculated
        ? `Calculated ${C.field}`
        : `${C.operation === "distinctCount" ? "Distinct count" : C.operation[0].toUpperCase() + C.operation.slice(1)} of ${C.field}`),
    I = (C) => (isNullOrEmptyValue(C) ? "(blank)" : String(C)),
    O = e.rowFields.length
      ? e.rowFields.map((C) => {
          const T = e.dateGroups?.find((R) => R.field === C);
          return T ? `${C} (${T.unit})` : C;
        })
      : ["Row labels"];
  for (const C of A)
    for (const T of s)
      O.push(
        [...C.parts.map(I), ...(C.subtotal ? ["Subtotal"] : []), F(T)].join(
          " / ",
        ),
      );
  if (e.grandTotals && a.length)
    for (const C of s) O.push(`Grand Total / ${F(C)}`);
  s.forEach((C, T) => {
    if (C.showAs && C.showAs !== "value")
      for (let N = O.length - 1; N >= Math.max(1, o.length); N--)
        (N - Math.max(1, o.length)) % s.length === T &&
          (O[N] =
            `${O[N]} (% of ${C.showAs === "percentGrand" ? "grand" : C.showAs === "percentRow" ? "row" : "column"} total)`);
  });
  const j = [],
    z = [...A, ...(e.grandTotals && a.length ? [null] : [])],
    V = [O],
    D = (C, T, N, R) => {
      const L = s[T].showAs;
      if (!L || L === "value" || C === null) return C;
      const $ = c(L === "percentGrand" ? y : L === "percentRow" ? N : R, T);
      return typeof C == "number" && typeof $ == "number" && $ !== 0
        ? C / $
        : "#DIV/0!";
    },
    B = [O],
    _ = [];
  for (const C of k) {
    const T = C.subtotal
        ? [
            ...C.parts,
            "Subtotal",
            ...Array(o.length - C.parts.length - 1).fill(null),
          ]
        : C.parts.length
          ? [...C.parts]
          : ["Total"],
      N = [...T];
    for (const R of z)
      s.forEach((L, $) => {
        const U = c(
          R ? v.get(JSON.stringify([C.key, R.key])) : g.get(C.key),
          $,
        );
        (N.push(U), T.push(D(U, $, g.get(C.key), R ? b.get(R.key) : y)));
      });
    (C.subtotal && _.push(B.length), B.push(T), V.push(N), j.push(C));
  }
  if (e.grandTotals && o.length) {
    const C = ["Grand Total", ...Array(o.length - 1).fill(null)],
      T = [...C];
    for (const N of z)
      s.forEach((R, L) => {
        const $ = c(N ? b.get(N.key) : y, L);
        (T.push($), C.push(D($, L, y, N ? b.get(N.key) : y)));
      });
    (B.push(C), V.push(T), j.push(null), _.push(B.length - 1));
  }
  return {
    matrix: B,
    baseMatrix: V,
    rowAxes: j,
    columnAxes: z,
    chartRows: k
      .filter((C) => !C.subtotal)
      .map((C) => ({
        label: C.parts.map(I).join(" / ") || "Total",
        values: s.map((T, N) => D(c(g.get(C.key), N), N, g.get(C.key), y)),
      })),
    subtotalRows: _,
    measures: s.map((C, T) => ({
      name: F(C),
      field: C.field,
      operation: C.operation,
      result: c(y, T),
      numericCount: y[T].numeric,
      nonblankCount: y[T].count,
    })),
  };
}
const pivotShowValuesAsModes = [
    "value",
    "percentGrand",
    "percentRow",
    "percentColumn",
  ],
  requirePivotFieldName = (e) =>
    typeof e == "string" && e.trim() && e.length <= 255
      ? e
      : throwErrorMessage("Choose a valid pivot field."),
  isIsoDateString = (e) =>
    typeof e == "string" &&
    /^\d{4}-\d{2}-\d{2}$/.test(e) &&
    !e.startsWith("0000") &&
    Number.isFinite(Date.parse(e)) &&
    new Date(e).toISOString().slice(0, 10) === e,
  requireBoundedList = (e, r) =>
    Array.isArray(e) && e.length <= r
      ? e
      : throwErrorMessage(
          `Pivot controls require a list of at most ${r} entries.`,
        ),
  buildSlicerItemKey = (e) =>
    JSON.stringify([
      typeof e,
      typeof e == "string" ? e.toLocaleLowerCase("en-US") : e,
    ]);
function normalizePivotControls(e) {
  const r = {};
  (e.slicers !== undefined &&
    (r.slicers = requireBoundedList(e.slicers, 20).map((n) => {
      if (!isNonArrayObject(n))
        return throwErrorMessage("Invalid pivot slicer.");
      const o = requireBoundedList(n.selected, 500).map(
        (a) => (
          a !== null &&
            typeof a != "boolean" &&
            !(typeof a == "number" && Number.isFinite(a)) &&
            !(typeof a == "string" && a.length <= 32767) &&
            throwErrorMessage("Invalid slicer item."),
          a
        ),
      );
      return (
        new Set(o.map(buildSlicerItemKey)).size !== o.length &&
          throwErrorMessage("Slicer selections must be unique."),
        { field: requirePivotFieldName(n.field), selected: o }
      );
    })),
    e.timelines !== undefined &&
      (r.timelines = requireBoundedList(e.timelines, 10).map((n) =>
        !isNonArrayObject(n) ||
        !isIsoDateString(n.from) ||
        !isIsoDateString(n.to) ||
        n.from > n.to ||
        !["year", "quarter", "month"].includes(String(n.unit)) ||
        !["iso", "1900", "1904"].includes(String(n.dateSystem))
          ? throwErrorMessage(
              "Timeline requires ordered YYYY-MM-DD bounds, a period and an explicit date system.",
            )
          : {
              field: requirePivotFieldName(n.field),
              from: n.from,
              to: n.to,
              unit: n.unit,
              dateSystem: n.dateSystem,
            },
      )),
    e.sorts !== undefined &&
      (r.sorts = requireBoundedList(e.sorts, 100).map((n) =>
        !isNonArrayObject(n) ||
        !["ascending", "descending"].includes(String(n.direction))
          ? throwErrorMessage(
              "Pivot field sorting requires ascending or descending order.",
            )
          : { field: requirePivotFieldName(n.field), direction: n.direction },
      )));
  for (const n of [r.slicers, r.timelines, r.sorts])
    n &&
      new Set(n.map((o) => o.field)).size !== n.length &&
      throwErrorMessage("Choose only one control of each kind per field.");
  if (e.chart !== undefined) {
    const n = e.chart;
    if (
      !isNonArrayObject(n) ||
      !["column", "bar", "line"].includes(String(n.kind)) ||
      !Number.isInteger(n.valueIndex) ||
      Number(n.valueIndex) < 0 ||
      Number(n.valueIndex) >= Number(e.values?.length)
    )
      return throwErrorMessage(
        "PivotChart requires a supported chart and an existing value field.",
      );
    r.chart = { kind: n.kind, valueIndex: n.valueIndex };
  }
  return r;
}
function validatePivotFieldReferences(e, r) {
  const t = new Set(e.calculatedFields?.map((a) => a.name)),
    n = [
      ...e.rowFields,
      ...e.columnFields,
      ...(e.filters ?? []).map((a) => a.field),
      ...(e.slicers ?? []).map((a) => a.field),
      ...(e.timelines ?? []).map((a) => a.field),
      ...(e.dateGroups ?? []).map((a) => a.field),
      ...(e.sorts ?? []).map((a) => a.field),
      ...e.values.filter((a) => !t.has(a.field)).map((a) => a.field),
      ...(e.calculatedFields ?? []).flatMap((a) => [
        a.field,
        ...(a.otherField ? [a.otherField] : []),
      ]),
    ];
  for (const a of n)
    r.includes(a) ||
      throwErrorMessage(
        `Broken pivot field “${a}”: source field was not found. Restore the header or remove/reselect this field before refreshing. Historical totals are not current.`,
      );
  for (const a of e.sorts ?? [])
    [...e.rowFields, ...e.columnFields].includes(a.field) ||
      throwErrorMessage(
        `Sort field “${a.field}” must be a row or column group.`,
      );
}
function normalizeTimelineDate(e, r, t, n) {
  if (e === null || e === "") return null;
  if (r === "iso")
    return isIsoDateString(e)
      ? e
      : throwErrorMessage(
          `Timeline field “${t}” has an invalid ISO date on source row ${n + 1}.`,
        );
  if (
    typeof e != "number" ||
    !Number.isFinite(e) ||
    e < (r === "1900" ? 1 : 0) ||
    e > 2958465
  )
    return throwErrorMessage(
      `Timeline field “${t}” requires ${r} date serials on source row ${n + 1}.`,
    );
  const o = Math.floor(e);
  if (r === "1900" && o === 60) return "1900-02-29";
  const a = new Date(
    (r === "1900" ? Date.UTC(1899, 11, 31) : Date.UTC(1904, 0, 1)) +
      (o - (r === "1900" && o >= 60 ? 1 : 0)) * 864e5,
  );
  return a.getUTCFullYear() > 9999
    ? throwErrorMessage("Timeline date exceeds year 9999.")
    : a.toISOString().slice(0, 10);
}
function filterPivotSourceRows(e, r, t) {
  validatePivotFieldReferences(e, r);
  const n = (e.slicers ?? []).map((a) => ({
      index: r.indexOf(a.field),
      keys: new Set(a.selected.map(buildSlicerItemKey)),
    })),
    o = (e.timelines ?? []).map((a) => ({ ...a, index: r.indexOf(a.field) }));
  return t.filter((a) => {
    const s = o.map((d) =>
      normalizeTimelineDate(a.values[d.index], d.dateSystem, d.field, a.index),
    );
    return (
      n.every((d) => d.keys.has(buildSlicerItemKey(a.values[d.index]))) &&
      o.every((d, u) => s[u] !== null && s[u] >= d.from && s[u] <= d.to)
    );
  });
}
function buildTimelinePeriods(e, r) {
  const t = new Map();
  for (const n of e) {
    const o = normalizeTimelineDate(n, r.dateSystem, r.field, 0);
    if (!o) continue;
    const a = Number(o.slice(0, 4)),
      s = Number(o.slice(5, 7)),
      d =
        r.unit === "year"
          ? 1
          : r.unit === "quarter"
            ? Math.floor((s - 1) / 3) * 3 + 1
            : s,
      u = r.unit === "year" ? 12 : r.unit === "quarter" ? d + 2 : d,
      l = String(a).padStart(4, "0"),
      c =
        r.unit === "year"
          ? l
          : r.unit === "quarter"
            ? `${l} Q${Math.ceil(s / 3)}`
            : `${l}-${String(s).padStart(2, "0")}`,
      h = new Date(0);
    (h.setUTCFullYear(a, u, 0),
      h.setUTCHours(0, 0, 0, 0),
      t.set(c, {
        label: c,
        from: `${l}-${String(d).padStart(2, "0")}-01`,
        to: h.toISOString().slice(0, 10),
      }));
  }
  return [...t.values()].sort((n, o) => n.from.localeCompare(o.from));
}
const isPlainDataObject = (e) =>
    !!e &&
    typeof e == "object" &&
    !Array.isArray(e) &&
    (Object.getPrototypeOf(e) === Object.prototype ||
      Object.getPrototypeOf(e) === null),
  isBoundedString = (e, r) =>
    typeof e == "string" && e.length > 0 && e.length <= r;
function normalizeRowProvenance(e, r) {
  const t = e.sheets;
  if (
    e.rowOrigins === undefined &&
    !t.some((a) => a.rowOriginIds !== undefined)
  )
    return;
  if (
    !Array.isArray(e.rowOrigins) ||
    e.rowOrigins.length > workbookLimits.maxCells
  )
    throw Error(
      "Row provenance requires a registry of at most 100,000 source rows.",
    );
  const n = new Set();
  r.rowOrigins = e.rowOrigins.map((a) => {
    if (
      !isPlainDataObject(a) ||
      !isBoundedString(a.id, 100) ||
      n.has(a.id) ||
      !isBoundedString(a.sourceId, 100) ||
      !isBoundedString(a.sourceName, 255) ||
      !isBoundedString(a.sourceSheet, 31) ||
      !["file", "snapshot"].includes(String(a.kind)) ||
      !Number.isInteger(a.originalRow) ||
      Number(a.originalRow) < 1 ||
      Number(a.originalRow) > workbookLimits.maxRows
    )
      throw Error("Invalid, duplicated or out-of-bounds original source row.");
    return (
      n.add(a.id),
      {
        id: a.id,
        sourceId: a.sourceId,
        sourceName: a.sourceName,
        sourceSheet: a.sourceSheet,
        originalRow: a.originalRow,
        kind: a.kind,
      }
    );
  });
  let o = 0;
  for (let a = 0; a < t.length; a++) {
    const s = t[a].rowOriginIds;
    if (s === undefined) continue;
    if (!isPlainDataObject(s))
      throw Error("Source row associations must be a plain index map.");
    const d = {};
    for (const [u, l] of Object.entries(s)) {
      if (
        !/^(0|[1-9]\d*)$/.test(u) ||
        Number(u) >= Number(t[a].rowCount) ||
        typeof l != "string" ||
        !n.has(l) ||
        ++o > workbookLimits.maxCells
      )
        throw Error("Invalid or over-capacity source row association.");
      d[u] = l;
    }
    r.sheets[a].rowOriginIds = d;
  }
}
function findRowOrigin(e, r, t) {
  const n = e.sheets.find((u) => u.id === r)?.rowOriginIds?.[t],
    o = n && e.rowOrigins?.find((u) => u.id === n);
  return o ? { ...o } : null;
}
function validateRowOrigins(e, r) {
  var n;
  if (!((n = e.rowOrigins) != null && n.length)) return;
  const t = new Map((r.rowOrigins ?? []).map((o) => [o.id, o]));
  for (const o of e.rowOrigins)
    if (
      !t.has(o.id) ||
      [
        "id",
        "sourceId",
        "sourceName",
        "sourceSheet",
        "originalRow",
        "kind",
      ].some((a) => t.get(o.id)[a] !== o[a])
    )
      throw Error(
        "Original source row identities cannot be changed or removed. Keep the source registry and change only current row associations.",
      );
}
const basicAggregateOperations = ["sum", "count", "average", "min", "max"],
  pivotAggregateOperations = [...basicAggregateOperations, "distinctCount"];

function assertUniqueTableNames(e) {
  const r = new Set(),
    t = new Set();
  for (const n of e.sheets)
    for (const o of n.tables ?? [])
      ((r.has(o.name.toLowerCase()) || t.has(o.id)) &&
        throwWorkbookError(
          "Table names and IDs must be unique across the workbook.",
        ),
        r.add(o.name.toLowerCase()),
        t.add(o.id));
}
function assertWithinCellBudget(e) {
  return (
    e.sheets.reduce((r, t) => r + Object.keys(t.cells).length, 0) >
      workbookLimits.maxCells &&
      throwWorkbookError(
        "The result would exceed the 100,000 populated-cell workbook budget.",
      ),
    e
  );
}
function findSourceFieldIndex(e, r) {
  const t = e.indexOf(r);
  return t >= 0
    ? t
    : throwWorkbookError(
        `Source field “${r}” was not found. Refresh the field selection.`,
      );
}
function buildUniqueSheetName(e, r) {
  const t =
    r
      .replace(/[\\/\[\]*?:]/g, " ")
      .trim()
      .slice(0, 31) || "Analysis";
  let n = t,
    o = 2;
  for (; e.sheets.some((a) => a.name.toLowerCase() === n.toLowerCase());)
    n = `${t.slice(0, 25)} (${o++})`;
  return n;
}
const computePivotRefreshFingerprint = (e) => {
  const r = JSON.stringify(e);
  r.length > 16e6 &&
    throwWorkbookError(
      "Pivot refresh evidence exceeds the 16-million-character fingerprint budget. Use a smaller source.",
    );
  const t = [];
  for (let n = 0; n < r.length; n += 5e5)
    t.push(hashPictureBytes(new Uint8Array(strToU8(r.slice(n, n + 5e5)))));
  return hashPictureBytes(
    new Uint8Array(strToU8(JSON.stringify([r.length, ...t]))),
  );
};
function findPivotDefinition(e, r) {
  const t = e.sheets.find((a) => {
      return a.pivots?.some((d) => d.id === r);
    }),
    n = t?.pivots?.find((a) => a.id === r);
  return (
    (!t || !n) && throwWorkbookError("Pivot definition not found."),
    { owner: t, definition: n }
  );
}
function buildPivotStatusRecord(e, r, t) {
  return {
    workbookId: e.id,
    revision: e.revision,
    pivotId: r.id,
    name: r.name,
    source: structuredClone(r.source),
    sourceRevision: r.sourceRevision,
    refreshedAt: r.refresh?.at ?? null,
    sourceFingerprint: r.refresh?.sourceFingerprint ?? null,
    status: t.state,
    diagnostic: t.reason,
    outputRange: r.outputRange,
    rowFields: [...r.rowFields],
    columnFields: [...r.columnFields],
    values: structuredClone(r.values),
    filters: structuredClone(r.filters ?? []),
    slicers: structuredClone(r.slicers ?? []),
    timelines: structuredClone(r.timelines ?? []),
    sorts: structuredClone(r.sorts ?? []),
    dateGroups: structuredClone(r.dateGroups ?? []),
    calculatedFields: structuredClone(r.calculatedFields ?? []),
    chart: structuredClone(r.chart ?? null),
    population: r.refresh
      ? {
          source: r.refresh.sourceRows,
          included: r.refresh.includedRows,
          excluded: r.refresh.excludedRows,
        }
      : null,
    measures: structuredClone(r.refresh?.measures ?? []),
    semantics:
      "Controls intersect. Slicer items use type-sensitive, case-insensitive equality; an empty selection includes no records. Hidden source rows are included. Percentages use filtered aggregate totals. Drill-through returns numerator records, including blank/nonnumeric measure inputs; it does not expand denominator populations.",
  };
}
const tableMutationLimits = {
  changedCells: 500,
  appendedRows: 100,
  formulaLength: 4096,
};
function validateCalculatedColumnFormula(e, r = false) {
  (typeof e != "string" ||
    e.length > tableMutationLimits.formulaLength ||
    e !== e.trim() ||
    !e.startsWith("=")) &&
    throwWorkbookError(
      "Calculated column formulas must begin with = and contain at most 4,096 characters, without outer whitespace.",
    );
  try {
    validateNameFormulaText(e);
  } catch (t) {
    throwWorkbookError(
      `Calculated column formula: ${t.message.replace(/^XLSX defined names: /, "").replace(/ Import\/export stopped.*$/, "")}`,
    );
  }
  return (
    !r &&
      /#REF!/.test(
        stripFormulaStringLiterals(e.replace(/"(?:[^"]|"")*"/g, '""')),
      ) &&
      throwWorkbookError(
        "Repair broken #REF! references before saving a calculated column rule.",
      ),
    e
  );
}
const a1CellPattern = "\\$?[A-Za-z]{1,3}\\$?[1-9]\\d{0,6}",
  a1ReferenceAnchorPattern = new RegExp(
    `^(${a1CellPattern}(?::${a1CellPattern})?)(?![A-Za-z0-9_.$:#\\[])`,
  );
function parseExternalLinkRange(e, r = 5e3) {
  const t = (p) => {
      const m = /^\$?([A-Za-z]{1,3})\$?([1-9]\d{0,6})$/.exec(p);
      if (!m)
        throw Error("External links require bounded A1 cells or rectangles.");
      let v = 0;
      for (const b of m[1].toUpperCase()) v = v * 26 + b.charCodeAt(0) - 64;
      const g = Number(m[2]);
      if (v > 16384 || g > 1048576)
        throw Error("External reference exceeds worksheet bounds.");
      return { row: g, col: v };
    },
    [n, o = n, a] = e.split(":");
  if (a) throw Error("Invalid external range.");
  const s = t(n),
    d = t(o),
    u = [],
    l = Math.min(s.row, d.row),
    c = Math.max(s.row, d.row),
    h = Math.min(s.col, d.col),
    f = Math.max(s.col, d.col);
  if ((c - l + 1) * (f - h + 1) > r)
    throw Error(
      `External link ranges are limited to ${r.toLocaleString()} cells.`,
    );
  for (let p = l; p <= c; p++)
    for (let m = h; m <= f; m++) {
      let v = m,
        g = "";
      for (; v;)
        ((g = String.fromCharCode(65 + ((v - 1) % 26)) + g),
          (v = Math.floor((v - 1) / 26)));
      u.push(g + p);
    }
  return u;
}
function tokenizeExternalReferences(e) {
  const r = [];
  for (let t = 0; t < e.length;) {
    if (e[t] === '"') {
      for (t++; t < e.length;)
        if (e[t++] === '"')
          if (e[t] === '"') t++;
          else break;
      continue;
    }
    const n = t,
      o = /^'((?:[^']|'')*)'\s*!/.exec(e.slice(t)),
      a = /^\[([^\[\]\r\n]+)\]([A-Za-z_][A-Za-z0-9_.]*)!/.exec(e.slice(t));
    let s,
      d = 0;
    if (
      (o
        ? ((s = o[1].replace(/''/g, "'")), (d = o[0].length))
        : a && ((s = `[${a[1]}]${a[2]}`), (d = a[0].length)),
      s)
    ) {
      const u = /^(.*\[[^\[\]\r\n]+\])([^\[\]:\r\n]+)$/.exec(s),
        l = a1ReferenceAnchorPattern.exec(e.slice(t + d));
      if (u && l && (!t || !/[A-Za-z0-9_.\]]/.test(e[t - 1])))
        try {
          (parseExternalLinkRange(l[1]),
            (t += d + l[0].length),
            r.push({
              start: n,
              end: t,
              token: e.slice(n, t),
              target: u[1],
              sheet: u[2],
              range: l[1].toUpperCase(),
            }));
          continue;
        } catch {}
      t += d;
      continue;
    }
    if (e[t] === "[") {
      let u = 0;
      do (e[t] === "[" && u++, e[t] === "]" && u--, t++);
      while (t < e.length && u > 0);
    } else if (e[t] === "'") {
      for (t++; t < e.length;)
        if (e[t++] === "'")
          if (e[t] === "'") t++;
          else break;
    } else t++;
  }
  return r;
}
function rewriteExternalReferences(e, r) {
  const t = tokenizeExternalReferences(e);
  for (const n of t.reverse()) {
    const o = r(n);
    o !== undefined && (e = e.slice(0, n.start) + o + e.slice(n.end));
  }
  return e;
}
const formatExternalSheetReference = (e, r, t) =>
    `'${(e + r).replace(/'/g, "''")}'!${t}`,
  workbookLinkLimits = Object.freeze({
    sources: 8,
    sheets: 32,
    cells: 5e3,
    formulas: 500,
    metadataCharacters: 2e6,
  }),
  workbookQualifierSchema = boundedTextSchema(2048).refine(
    (e) => /^[^\[\]]*\[[^\[\]]+\]$/.test(e),
    "Use an exact workbook qualifier such as [Source.xlsx].",
  ),
  tU = string().regex(/^[a-f0-9]{64}$/),
  linkedCellValueSchema = discriminatedUnion("type", [
    strictObject({ type: literal("number"), value: number().finite() }),
    strictObject({ type: literal("text"), value: string().max(32767) }),
    strictObject({ type: literal("boolean"), value: boolean() }),
    strictObject({ type: literal("blank"), value: _null() }),
    strictObject({
      type: literal("error"),
      value: string()
        .regex(/^#[A-Z0-9/?!_.]+$/)
        .max(30),
    }),
  ]),
  workbookLinkSourceSchema = strictObject({
    target: workbookQualifierSchema,
    sourceId: boundedTextSchema(100),
    sourceName: boundedTextSchema(200),
    sourceRevision: number().int().nonnegative(),
    sourceFingerprint: tU,
    resolvedAt: isoDateTimeSchemaFactory(),
    sheets: array(
      strictObject({
        id: boundedTextSchema(100),
        name: boundedTextSchema(31),
        cells: record(
          string().regex(/^[A-Z]{1,3}[1-9]\d{0,6}$/),
          linkedCellValueSchema,
        ),
      }),
    )
      .min(1)
      .max(32),
  }),
  workbookLinkPackageSchema = strictObject({
    schemaVersion: literal(1),
    sources: array(workbookLinkSourceSchema).max(8),
    packageTargets: array(
      strictObject({
        key: workbookQualifierSchema,
        target: boundedTextSchema(4096),
        part: boundedTextSchema(1024),
        sourceSha256: tU,
      }),
    )
      .max(100)
      .optional(),
  }),
  validatedWorkbookLinkRegistry = new WeakSet();
function attachWorkbookLinkMetadata(e, r) {
  if (e.workbookLinks === undefined) return;
  if (
    e.workbookLinks &&
    typeof e.workbookLinks == "object" &&
    validatedWorkbookLinkRegistry.has(e.workbookLinks)
  ) {
    r.workbookLinks = e.workbookLinks;
    return;
  }
  if (
    JSON.stringify(e.workbookLinks).length >
    workbookLinkLimits.metadataCharacters
  )
    throw Error("Workbook link metadata exceeds 2 million characters.");
  const t = workbookLinkPackageSchema.parse(e.workbookLinks);
  if (new Set(t.sources.map((d) => d.target)).size !== t.sources.length)
    throw Error("Workbook link targets must be unique.");
  if (
    new Set(t.packageTargets?.map((d) => d.key)).size !==
    (t.packageTargets?.length ?? 0)
  )
    throw Error("Package link identities must be unique.");
  let n = 0,
    o = 0;
  for (const d of t.sources) {
    if (
      new Set(d.sheets.map((u) => u.name.toLowerCase())).size !==
      d.sheets.length
    )
      throw Error("Linked worksheet names must be unique.");
    for (const u of d.sheets) {
      o++;
      for (const l of Object.keys(u.cells)) (parseExternalLinkRange(l), n++);
    }
  }
  if (n > workbookLinkLimits.cells || o > workbookLinkLimits.sheets)
    throw Error("Workbook link cache exceeds its cell or sheet limit.");
  ((r.workbookLinks = deepFreezeValue(t)),
    validatedWorkbookLinkRegistry.add(r.workbookLinks));
}
function encodeWorkbookFingerprint(e) {
  const r = new TextEncoder().encode(
    JSON.stringify([
      e.id,
      e.name,
      e.revision,
      e.calculation,
      e.datePolicy,
      e.definedNames,
      e.workbookLinks,
      e.sheets.map((t) => [
        t.id,
        t.name,
        t.rowCount,
        t.colCount,
        t.hidden,
        t.veryHidden,
        t.hiddenRows,
        t.hiddenCols,
        t.tables,
        t.merges,
        t.outlineGroups,
        Object.entries(t.cells).map(([n, o]) => [n, o.raw, !!o.literal]),
      ]),
    ]),
  );
  return sha256HexDigest(r);
}
function buildExternalLinkAliases(e) {
  const r = [],
    t = new Map(),
    n = new Map(),
    o = new Set(e.sheets.map((d) => d.name.toLowerCase()));
  for (const [d, u] of (e.workbookLinks?.sources ?? []).entries())
    for (const [l, c] of u.sheets.entries()) {
      let h = `_LG_LINK_${d}_${l}`;
      for (; o.has(h.toLowerCase());) h += "_";
      o.add(h.toLowerCase());
      let f = `linked:${d}:${l}`;
      for (; e.sheets.some((m) => m.id === f);) f += "_";
      const p = {};
      for (const [m, v] of Object.entries(c.cells))
        ((p[m] = { raw: v.value, literal: true }), t.set(`${f}:${m}`, v));
      (r.push({ id: f, name: h, rowCount: 1048576, colCount: 16384, cells: p }),
        n.set(JSON.stringify([u.target, c.name.toLowerCase()]), h));
    }
  return {
    sheets: r,
    values: t,
    rewrite: (d) =>
      rewriteExternalReferences(d, (u) => {
        const l = e.workbookLinks?.sources.find((f) => f.target === u.target),
          c = l?.sheets.find(
            (f) => f.name.toLowerCase() === u.sheet.toLowerCase(),
          );
        if (!(
          !c ||
          !parseExternalLinkRange(u.range).every((f) =>
            Object.hasOwn(c.cells, f),
          )
        ))
          return `'${n.get(JSON.stringify([u.target, u.sheet.toLowerCase()]))}'!${u.range}`;
      }),
  };
}
const databaseFunctionNames = [
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
  notDatabaseFunctionMarker = Symbol("not a database function"),
  databaseFunctionLimits = {
    databaseCells: 2e4,
    criteriaCells: 1e3,
    predicateComparisons: 1e5,
    textComparisonSteps: 5e6,
    criterionCharacters: 255,
    headerCharacters: 255,
    formulaEvaluations: 1e4,
  };
function evaluateDatabaseFunction(e, r, t) {
  if (!databaseFunctionNames.includes(e)) return notDatabaseFunctionMarker;
  const n = (T = "#VALUE!") => {
    throw t.issue(T);
  };
  if (r.length !== 3 || r[0] == null || r[2] == null) return n();
  const o = (T, N) => {
      const R = t.retrieve(T);
      if (t.isIssue(R)) throw R;
      if (!Array.isArray(R) || R.length < 2) return n();
      if (R.length > N) return n("#LIMIT!");
      if (!Array.isArray(R[0])) return n();
      const L = R,
        $ = L[0].length;
      if (!$) return n();
      if (L.length * $ > N) return n("#LIMIT!");
      for (let U = 0; U < L.length; U++)
        if (!Array.isArray(L[U]) || L[U].length !== $) return n();
      for (const U of L)
        for (const G of U)
          if (t.isIssue(G)) {
            if (
              [
                "#LIMIT!",
                "#CYCLE!",
                "#BLOCKED!",
                "#TIMEOUT!",
                "#ITER!",
                "#ERROR!",
              ].includes(G.error)
            )
              throw G;
          } else {
            if (
              G != null &&
              !["number", "string", "boolean"].includes(typeof G)
            )
              return n();
            if (typeof G == "number" && !Number.isFinite(G)) return n("#NUM!");
          }
      return L;
    },
    a = o(r[0], databaseFunctionLimits.databaseCells),
    s = o(r[2], databaseFunctionLimits.criteriaCells);
  if (
    (a.length - 1) * (s.length - 1) * s[0].length >
    databaseFunctionLimits.predicateComparisons
  )
    return n("#LIMIT!");
  const d = (T) => {
      if (t.isIssue(T)) throw T;
      return typeof T != "string" ||
        !T.trim() ||
        T.length > databaseFunctionLimits.headerCharacters
        ? n()
        : T.toLowerCase();
    },
    u = a[0].map(d);
  if (new Set(u).size !== u.length) return n();
  const l = s[0].map((T) => {
      if (t.isIssue(T)) throw T;
      return T == null || T === "" ? -1 : u.indexOf(d(T));
    }),
    c = e === "DCOUNT" || e === "DCOUNTA";
  let h;
  if (r[1] == null) {
    if (!c) return n();
  } else {
    const T = t.retrieve(r[1]);
    if (t.isIssue(T)) throw T;
    if (typeof T == "number") {
      if (!Number.isInteger(T) || T < 1 || T > u.length) return n();
      h = T - 1;
    } else if (typeof T == "string" && T.length) {
      if (((h = u.indexOf(T.toLowerCase())), h < 0)) return n();
    } else return n();
  }
  let f = 0;
  const p = (T) => {
      ((f += T),
        f > databaseFunctionLimits.textComparisonSteps && n("#LIMIT!"));
    },
    m = (T, N, R) =>
      R === "<>"
        ? T !== N
        : R === "<"
          ? T < N
          : R === ">"
            ? T > N
            : R === "<="
              ? T <= N
              : R === ">="
                ? T >= N
                : T === N,
    v = (T) => {
      if (t.isIssue(T)) throw T;
      if (T == null || T === "") return null;
      if (typeof T == "number" || typeof T == "boolean") return (Q) => Q === T;
      if (typeof T != "string") return n();
      if (T.length > databaseFunctionLimits.criterionCharacters)
        return n("#LIMIT!");
      const N = /^(<=|>=|<>|=|<|>)?(.*)$/s.exec(T),
        R = N[1] ?? "=",
        L = N[2];
      if (N[1] && /^[<>=]/.test(L)) return n();
      if (!L && N[1])
        return R !== "=" && R !== "<>"
          ? n()
          : (Q) => (R === "=" ? Q == null || Q === "" : Q != null && Q !== "");
      if (/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(L.trim())) {
        const Q = Number(L);
        return Number.isFinite(Q)
          ? (X) => (typeof X == "number" ? m(X, Q, R) : R === "<>")
          : n("#NUM!");
      }
      if (/^(TRUE|FALSE)$/i.test(L)) {
        if (R !== "=" && R !== "<>") return n();
        const Q = L.toUpperCase() === "TRUE";
        return (X) => (R === "=" ? X === Q : X !== Q);
      }
      const $ = L.toLowerCase(),
        U = [];
      let G = false;
      for (let Q = 0; Q < $.length; Q++) {
        const X = $[Q];
        if (X === "~") {
          const J = $[Q + 1];
          J && ["*", "?", "~"].includes(J)
            ? (U.push({ kind: "literal", char: J }), Q++, (G = true))
            : U.push({ kind: "literal", char: X });
        } else
          X === "*" || X === "?"
            ? ((G = true),
              (X === "?" || U.at(-1)?.kind !== "many") &&
                U.push({ kind: X === "*" ? "many" : "one" }))
            : U.push({ kind: "literal", char: X });
      }
      return ["<", ">", "<=", ">="].includes(R)
        ? G
          ? n()
          : (Q) =>
              typeof Q != "string"
                ? false
                : (p(Q.length + $.length), m(Q.toLowerCase(), $, R))
        : (!N[1] && !G && U.push({ kind: "many" }),
          (Q) => {
            if (typeof Q != "string") return R === "<>";
            p((Q.length + 1) * (U.length + 1));
            const X = Q.toLowerCase();
            let J = new Uint8Array(X.length + 1);
            J[0] = 1;
            for (const W of U) {
              const K = new Uint8Array(X.length + 1);
              if (W.kind === "many") {
                K[0] = J[0];
                for (let te = 1; te <= X.length; te++)
                  K[te] = J[te] || K[te - 1];
              } else
                for (let te = 1; te <= X.length; te++)
                  K[te] =
                    J[te - 1] && (W.kind === "one" || W.char === X[te - 1])
                      ? 1
                      : 0;
              J = K;
            }
            const q = !!J[X.length];
            return R === "<>" ? !q : q;
          });
    };
  let g = 0;
  const b = s.slice(1).map((T, N) =>
    T.map((R, L) => {
      return l[L] < 0 ? t.criterionFormula?.call(t, r[2], N + 1, L) : undefined;
    }),
  );
  l.forEach((T, N) => {
    T < 0 && !b.some((R) => R[N]) && n();
  });
  const y = s.slice(1).map((T, N) =>
    T.map((R, L) => {
      if (l[L] >= 0) return v(R);
      if (!b[N][L] && !(R == null || R === "")) return n();
    }),
  );
  b.flat().filter(Boolean).length * (a.length - 1) >
    databaseFunctionLimits.formulaEvaluations && n("#LIMIT!");
  let E = 0,
    A = 0,
    k = 0,
    F = 0,
    I = 1,
    O = 1 / 0,
    j = -1 / 0,
    z = 0,
    V = 0,
    D;
  for (const [T, N] of a.slice(1).entries()) {
    let R = false;
    for (const [U, G] of y.entries()) {
      let H = true;
      for (let Q = 0; Q < G.length; Q++) {
        const X = b[U][Q];
        if (X) {
          ++g > databaseFunctionLimits.formulaEvaluations && n("#LIMIT!");
          const W = X(T);
          if (t.isIssue(W)) throw W;
          (typeof W != "boolean" && n(), W || (H = false));
          continue;
        }
        const J = G[Q];
        if (!J) continue;
        const q = N[l[Q]];
        if (t.isIssue(q)) throw q;
        J(q) || (H = false);
      }
      H && (R = true);
    }
    if (!R || (E++, h === undefined)) continue;
    const L = N[h];
    if (((D = L), L != null && k++, t.isIssue(L))) {
      if (!["DCOUNT", "DCOUNTA", "DGET"].includes(e)) throw L;
      continue;
    }
    if (typeof L != "number") continue;
    (A++, (F += L), (I *= L), (O = Math.min(O, L)), (j = Math.max(j, L)));
    const $ = L - z;
    ((z += $ / A), (V += $ * (L - z)));
  }
  if (e === "DGET") {
    if (E !== 1) return n(E ? "#NUM!" : "#VALUE!");
    if (t.isIssue(D)) throw D;
    return D ?? 0;
  }
  if (e === "DCOUNT") return h === undefined ? E : A;
  if (e === "DCOUNTA") return h === undefined ? E : k;
  const B = (T) => (Number.isFinite(T) ? T : n("#NUM!"));
  if (e === "DSUM") return B(F);
  if (e === "DMAX") return A ? j : 0;
  if (e === "DMIN") return A ? O : 0;
  if (e === "DPRODUCT") return A ? B(I) : 0;
  if (e === "DAVERAGE") return A ? B(F / A) : n("#DIV/0!");
  const _ = e === "DSTDEV" || e === "DVAR";
  if (A < (_ ? 2 : 1)) return n("#DIV/0!");
  const C = Math.max(0, V / (A - (_ ? 1 : 0)));
  return B(e === "DSTDEV" || e === "DSTDEVP" ? Math.sqrt(C) : C);
}
const referenceSensitiveFunctionNames = [
    "SUBTOTAL",
    "AGGREGATE",
    "ISFORMULA",
    "FORMULATEXT",
  ],
  notReferenceFunctionMarker = Symbol("not-reference-function"),
  isFormulaCellValue = (e) =>
    typeof e?.raw == "string" && e.raw.startsWith("=") && !e.literal;
function containsSubtotalCall(e) {
  return /(?:^|[^A-Za-z0-9_.])(?:_xlfn\.(?:_xlws\.)?)?SUBTOTAL\s*\(/i.test(
    e.replace(/"(?:[^"]|"")*"|'(?:[^']|'')*'(?=\s*!)/g, ""),
  );
}
function containsSubtotalOrAggregateCall(e) {
  return /(?:^|[^A-Za-z0-9_.])(?:_xlfn\.(?:_xlws\.)?)?(?:SUBTOTAL|AGGREGATE)\s*\(/i.test(
    e.replace(/"(?:[^"]|"")*"|'(?:[^']|'')*'(?=\s*!)/g, ""),
  );
}
const aggregateFunctionNames = [
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
  builtinFunctionNames = [
    ...new Set([
      ...aggregateFunctionNames,
      ...supportedFunctionNames,
      ...modernFunctionNames,
      ...supportedFinancialFunctionNames,
      ...referenceSensitiveFunctionNames,
      ...databaseFunctionNames,
      ...statisticalFunctionNames,
      ...ownedLookupFunctionNames,
      ...localBindingFunctionNames,
      ...specialistFunctionNames,
      ...XL,
      ...significanceTestFunctionNames,
    ]),
  ].sort(),
  builtinFunctionNameSet = new Set(builtinFunctionNames),
  FormulaErrorClass = FormulaParser.FormulaError,
  createFormulaError = (e) => new FormulaErrorClass(e),
  formulaErrorFactory = createFormulaError,
  isFormulaError = (e) => e instanceof FormulaErrorClass,
  findFormulaErrorArg = (e) =>
    e.flatMap(flattenNestedValues).find(isFormulaError);
function findScopedDefinedName(e, r, t) {
  if (
    e.cell ||
    (e.call &&
      !e.qualifier &&
      builtinFunctionNames.includes(normalizeFunctionName(e.name)))
  )
    return;
  const n = e.qualifier
    ? r.sheets.find((d) => d.name.toLowerCase() === e.qualifier.toLowerCase())
        ?.id
    : t;
  if (n)
    return (
      r.definedNames?.find(
        (d) =>
          d.scopeSheetId === n && d.name.toLowerCase() === e.name.toLowerCase(),
      ) ??
      r.definedNames?.find(
        (d) => !d.scopeSheetId && d.name.toLowerCase() === e.name.toLowerCase(),
      )
    );
}
function resolveDefinedNameValue(e, r, t, n = [], o) {
  if (n.length > 32) throw new Error("#LIMIT!");
  const a = replaceTextSpans(e, m5(e), (s) => {
    if (s.call) return s.text;
    const d = findScopedDefinedName(s, r, t);
    if (!d || parseFormulaTree(d.refersTo)) return s.text;
    const u = definedNameScopeKey(d);
    if (n.includes(u)) throw new Error("#CYCLE!");
    const l =
      d.scopeSheetId ??
      (s.qualifier
        ? r.sheets.find(
            (h) => h.name.toLowerCase() === s.qualifier.toLowerCase(),
          ).id
        : t);
    if (o && !o(l)) throw new Error("#BLOCKED!");
    let c = resolveDefinedNameValue(d.refersTo.slice(1), r, l, [...n, u], o);
    if (l !== t) {
      const h = r.sheets.find((p) => p.id === l),
        f = m5(c);
      c = replaceTextSpans(c, f, (p) => {
        const m = f.indexOf(p),
          v = f[m - 1],
          g = v?.cell && c.slice(v.end, p.start).trim() === ":";
        return p.cell && !p.qualifier && !g
          ? formatQuotedSheetPrefix(h.name) + p.text
          : p.text;
      });
    }
    return `(${c})`;
  });
  if (a.length > 16384) throw new Error("#LIMIT!");
  return a;
}

export {
  databaseFunctionNames,
  formatExternalSheetReference,
  assertDrawingPreviewCurrent,
  buildSlicerItemKey,
  bucketPivotDateValue,
  encodeWorkbookFingerprint,
  F2,
  pivotAggregateOperations,
  builtinFunctionNames,
  containsSubtotalCall,
  tableMutationLimits,
  findFormulaErrorArg,
  builtinFunctionNameSet,
  attachWorkbookLinkMetadata,
  computePivotTable,
  inspectPictureDataUrl,
  evaluateDatabaseFunction,
  findPivotDefinition,
  parseExternalLinkRange,
  findRowOrigin,
  assertUniqueTableNames,
  computePivotRefreshFingerprint,
  workbookQualifierSchema,
  findSourceFieldIndex,
  validatePictureEvidence,
  containsSubtotalOrAggregateCall,
  normalizeRowProvenance,
  basicAggregateOperations,
  findScopedDefinedName,
  tokenizeExternalReferences,
  createPictureEvidence,
  referenceSensitiveFunctionNames,
  normalizePivotControls,
  filterPivotSourceRows,
  formulaErrorFactory,
  createFormulaError,
  notDatabaseFunctionMarker,
  pivotShowValuesAsModes,
  notReferenceFunctionMarker,
  buildTimelinePeriods,
  resolveDefinedNameValue,
  buildUniqueSheetName,
  buildPivotStatusRecord,
  buildExternalLinkAliases,
  validateRowOrigins,
  validatePivotRefreshEvidence,
  validatePivotFieldReferences,
  workbookLinkLimits,
  validateCalculatedColumnFormula,
  isFormulaError,
  assertWithinCellBudget,
  rewriteExternalReferences,
  isFormulaCellValue,
};
export { throwErrorMessage, isPlainDataObject, isFiniteNumber };
