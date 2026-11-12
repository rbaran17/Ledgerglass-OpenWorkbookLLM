import {
  isNonArrayObject,
  throwWorkbookError,
} from "./01-cell-styles-and-themes.js";
import { isNullOrEmptyValue } from "./03-defined-names.js";
import { boundedTextSchema } from "../01-schemas-date-functions-and-validation.js";
import { flattenNestedValues } from "../02-text-and-financial-functions.js";
import {
  workbookLimits,
  supportedFinancialFunctionNames,
} from "../02-text-and-financial-functions.js";
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
} from "../01-schemas-date-functions-and-validation.js";
import { assertSheetUnprotected } from "./01-cell-styles-and-themes.js";
import {
  FormulaParser,
  statisticalFunctionNames,
  XL,
  specialistFunctionNames,
  significanceTestFunctionNames,
  ownedLookupFunctionNames,
  modernFunctionNames,
} from "./02-formula-function-evaluators.js";
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

function inspectPictureDataUrl(dataURL) {
  if (
    typeof dataURL != "string" ||
    dataURL.length > Math.ceil(maxPictureHashBytes / 3) * 4 + 30
  )
    throw Error("Pictures must be local PNG/JPEG data of at most 2 MB.");
  const r = /^data:image\/(png|jpeg);base64,([A-Za-z0-9+/]+={0,2})$/.exec(
    dataURL,
  );
  if (!r || r[2].length % 4 !== 0)
    throw Error("Pictures require canonical local PNG/JPEG base64 data.");
  let bytes;
  try {
    const d = atob(r[2]);
    if (btoa(d) !== r[2]) throw Error();
    bytes = Uint8Array.from(d, (u) => u.charCodeAt(0));
  } catch {
    throw Error("The picture data is invalid.");
  }
  if (bytes.length > maxPictureHashBytes)
    throw Error("Pictures must be at most 2 MB.");
  let width = 0,
    height = 0;
  const a = () => {
    if (
      !width ||
      !height ||
      width > 1e4 ||
      height > 1e4 ||
      width * height > 4e7
    )
      throw Error(
        "Picture dimensions are invalid or exceed 40 megapixels / 10,000 pixels per side.",
      );
  };
  if (r[1] === "png") {
    if (
      bytes.length < 45 ||
      ![137, 80, 78, 71, 13, 10, 26, 10].every((p, m) => bytes[m] === p)
    )
      throw Error("The file contents do not match a PNG picture.");
    const d = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    if (
      d.getUint32(8) !== 13 ||
      String.fromCharCode(...bytes.subarray(12, 16)) !== "IHDR"
    )
      throw Error("PNG requires an intact first IHDR chunk.");
    if (
      ((width = d.getUint32(16)),
      (height = d.getUint32(20)),
      a(),
      !{
        0: [1, 2, 4, 8, 16],
        2: [8, 16],
        3: [1, 2, 4, 8],
        4: [8, 16],
        6: [8, 16],
      }[bytes[25]]?.includes(bytes[24]) ||
        bytes[26] !== 0 ||
        bytes[27] !== 0 ||
        bytes[28] > 1)
    )
      throw Error("Unsupported PNG pixel encoding.");
    let l = 8,
      c = 0,
      h = false,
      f = false;
    for (; l < bytes.length;) {
      if (++c > 4096 || l + 12 > bytes.length)
        throw Error("PNG chunk structure is invalid or excessive.");
      const p = d.getUint32(l),
        m = String.fromCharCode(...bytes.subarray(l + 4, l + 8));
      if (p > bytes.length - l - 12 || !/^[A-Za-z]{4}$/.test(m))
        throw Error("PNG chunk length or type is invalid.");
      if (m === "IHDR" && l !== 8) throw Error("PNG has duplicate dimensions.");
      if ((m === "IDAT" && p > 0 && (h = true), m === "IEND")) {
        if (p !== 0 || l + 12 !== bytes.length)
          throw Error("PNG has invalid trailing data.");
        f = true;
      }
      l += p + 12;
    }
    if (!h || !f) throw Error("PNG data or end marker is missing.");
  } else {
    if (bytes.length < 12 || bytes[0] !== 255 || bytes[1] !== 216)
      throw Error("The file contents do not match a JPEG picture.");
    let d = 2,
      u = 0,
      l = false,
      c = false;
    for (; d < bytes.length;) {
      if (++u > 65536) throw Error("JPEG has too many segments.");
      if (l) for (; d < bytes.length && bytes[d] !== 255;) d++;
      if (d >= bytes.length || bytes[d++] !== 255)
        throw Error("JPEG marker structure is invalid.");
      for (; bytes[d] === 255;) d++;
      const h = bytes[d++];
      if (l && (h === 0 || (h >= 208 && h <= 215))) continue;
      if (h === 217) {
        if (d !== bytes.length) throw Error("JPEG has trailing data.");
        c = true;
        break;
      }
      if (h === 1) continue;
      if (d + 2 > bytes.length) throw Error("JPEG segment is truncated.");
      const f = bytes[d] * 256 + bytes[d + 1];
      if (f < 2 || d + f > bytes.length)
        throw Error("JPEG segment length is invalid.");
      if (h >= 192 && h <= 207 && ![196, 200, 204].includes(h)) {
        if (![192, 193, 194].includes(h) || f < 8 || bytes[d + 2] !== 8)
          throw Error("Unsupported JPEG pixel encoding.");
        const p = bytes[d + 7];
        if (p < 1 || p > 4 || f !== 8 + p * 3)
          throw Error("JPEG frame components are invalid.");
        const m = bytes[d + 3] * 256 + bytes[d + 4],
          v = bytes[d + 5] * 256 + bytes[d + 6];
        if (width && (width !== v || height !== m))
          throw Error("JPEG dimensions conflict.");
        ((width = v), (height = m), a());
      }
      if (h === 218) {
        if (!width || f < 6 || f !== 6 + bytes[d + 2] * 2)
          throw Error("JPEG scan is invalid.");
        l = true;
      }
      d += f;
    }
    if (!l || !c) throw Error("JPEG scan or end marker is missing.");
    a();
  }
  return { bytes: bytes, extension: r[1], width: width, height: height };
}
function assertPlainTextLabel(e, r, t) {
  const n = validateDrawingText(e, r, t);
  if (!n.trim() || /[\u0000-\u001f\u007f]/.test(n))
    throw Error(
      `Invalid picture ${t}: use nonblank plain text without control characters.`,
    );
  return n;
}
function validatePictureEvidence(evidence, bytes) {
  if (
    !isNonArrayObject(evidence) ||
    ![Object.prototype, null].includes(Object.getPrototypeOf(evidence)) ||
    Reflect.ownKeys(evidence).some(
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
    evidence.version !== 1 ||
    typeof evidence.origin != "string" ||
    !["local-file", "xlsx-media", "legacy-native"].includes(evidence.origin)
  )
    throw Error("Invalid picture evidence origin.");
  const sha256 = hashPictureBytes(bytes);
  if (
    typeof evidence.sha256 != "string" ||
    !/^[a-f0-9]{64}$/.test(evidence.sha256) ||
    evidence.sha256 !== sha256 ||
    evidence.sourceReference !== `sha256:${sha256}` ||
    evidence.byteLength !== bytes.length
  )
    throw Error(
      "Picture evidence digest or byte length does not match the embedded image.",
    );
  const n = {
    version: 1,
    origin: evidence.origin,
    sourceReference: `sha256:${sha256}`,
    sha256: sha256,
    byteLength: bytes.length,
  };
  for (const o of [
    "filename",
    "sourceLocator",
    "containerFilename",
    "mediaReference",
  ])
    evidence[o] !== undefined &&
      (n[o] = assertPlainTextLabel(
        evidence[o],
        o === "sourceLocator" ? 2e3 : 512,
        o,
      ));
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
function createPictureEvidence(dataURL, r) {
  const { bytes } = inspectPictureDataUrl(dataURL),
    sha256 = hashPictureBytes(bytes);
  return validatePictureEvidence(
    {
      ...r,
      version: 1,
      sourceReference: `sha256:${sha256}`,
      sha256: sha256,
      byteLength: bytes.length,
    },
    bytes,
  );
}
function F2(workbook, sheetId, selection) {
  if (!workbook.sheets.some((sheet) => sheet.id === sheetId))
    throw Error("The worksheet no longer exists.");
  return {
    workbookId: workbook.id,
    revision: workbook.revision,
    sheetId: sheetId,
    ...(selection !== undefined ? { selection: selection } : {}),
    guard: JSON.stringify(workbook),
  };
}
function assertDrawingPreviewCurrent(workbook, sheetId, t = {}) {
  if (t.readOnly) throw Error("This workbook is read-only.");
  const sheet = workbook.sheets.find((sheet) => sheet.id === sheetId);
  if (!sheet) throw Error("The worksheet no longer exists.");
  assertSheetUnprotected(sheet);
  const snapshot = t.snapshot;
  if (
    snapshot &&
    (snapshot.workbookId !== workbook.id ||
      snapshot.revision !== workbook.revision ||
      snapshot.sheetId !== sheetId ||
      snapshot.guard !== JSON.stringify(workbook) ||
      snapshot.selection !== t.selection)
  )
    throw Error(
      "The drawing preview is stale. Reopen it or select the object again.",
    );
}
const throwErrorMessage = (e) => {
    throw Error(e);
  },
  isFiniteNumber = (e) => typeof e == "number" && Number.isFinite(e);
function validatePivotRefreshEvidence(refresh) {
  isNonArrayObject(refresh) ||
    throwErrorMessage("Invalid pivot refresh evidence.");
  const r = refresh;
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
  const measures = r.measures.map((measure) => {
    (!isNonArrayObject(measure) ||
      typeof measure.name != "string" ||
      !measure.name ||
      measure.name.length > 512 ||
      typeof measure.field != "string" ||
      !measure.field ||
      measure.field.length > 255 ||
      !["sum", "count", "average", "min", "max", "distinctCount"].includes(
        String(measure.operation),
      ) ||
      !(
        measure.result === null ||
        typeof measure.result == "boolean" ||
        isFiniteNumber(measure.result) ||
        (typeof measure.result == "string" && measure.result.length <= 32767)
      )) &&
      throwErrorMessage("Invalid pivot measure evidence.");
    const o = measure;
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
    measures: measures,
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
function SO(right, r) {
  (isNullOrEmptyValue(r) ||
    (right.count++,
    right.distinct?.add(
      JSON.stringify([
        typeof r,
        typeof r == "string" ? r.toLocaleLowerCase("en-US") : r,
      ]),
    )),
    isFiniteNumber(r) &&
      (right.numeric++,
      (right.sum += r),
      (right.min = Math.min(right.min, r)),
      (right.max = Math.max(right.max, r)),
      Number.isFinite(right.sum) ||
        throwErrorMessage("Aggregation exceeds finite numeric bounds.")));
}
function finalizeAggregateState(e, operation) {
  return operation === "count"
    ? e.count
    : operation === "distinctCount"
      ? (e.distinct?.size ?? 0)
      : operation === "sum"
        ? e.sum
        : e.numeric
          ? operation === "average"
            ? e.sum / e.numeric
            : operation === "min"
              ? e.min
              : e.max
          : null;
}
function computePivotTable(pivotDefinition, fields, rows) {
  const n = (C) => {
      const T = fields.indexOf(C);
      return T < 0
        ? throwErrorMessage(
            `Source field “${C}” was not found. Refresh the field selection.`,
          )
        : T;
    },
    o = pivotDefinition.rowFields.map(n),
    a = pivotDefinition.columnFields.map(n);
  for (const C of pivotDefinition.dateGroups ?? [])
    [...pivotDefinition.rowFields, ...pivotDefinition.columnFields].includes(
      C.field,
    ) ||
      throwErrorMessage(
        `Date group “${C.field}” must be selected as a row or column field.`,
      );
  for (const C of pivotDefinition.calculatedFields ?? [])
    (fields.some((field) => field.toLowerCase() === C.name.toLowerCase()) &&
      throwErrorMessage(
        `Calculated field “${C.name}” duplicates a source header.`,
      ),
      n(C.field),
      C.otherField && n(C.otherField));
  const s = pivotDefinition.values.map((value) => {
      const calculated = pivotDefinition.calculatedFields?.find(
        (calculatedField) => calculatedField.name === value.field,
      );
      return (
        calculated &&
          value.operation !== "sum" &&
          throwErrorMessage(
            `Calculated field “${value.field}” requires sum aggregation; its arithmetic is applied to source-field sums.`,
          ),
        {
          ...value,
          index: n(calculated?.field ?? value.field),
          calculated: calculated,
          other: calculated?.otherField ? n(calculated.otherField) : undefined,
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
  rows.length *
    s.length *
    Math.max(1, pivotDefinition.subtotals ? o.length : 1) *
    Math.max(1, pivotDefinition.subtotals ? a.length : 1) >
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
      const sum = R.sum,
        $ = N.other !== undefined ? R.right.sum : N.calculated.constant;
      N.calculated.operation === "divide" &&
        $ === 0 &&
        throwErrorMessage(
          `Calculated field “${N.field}” divides by zero in a group or total. Filter, clean or change the calculation; no result was written.`,
        );
      const U =
        N.calculated.operation === "add"
          ? sum + $
          : N.calculated.operation === "subtract"
            ? sum - $
            : N.calculated.operation === "multiply"
              ? sum * $
              : sum / $;
      return Number.isFinite(U)
        ? U
        : throwErrorMessage(
            `Calculated field “${N.field}” exceeds finite numeric bounds.`,
          );
    },
    h = (C, T) =>
      C.map((N) => {
        const R = pivotDefinition.dateGroups?.find(
            (dateGroup) => dateGroup.field === fields[N],
          ),
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
        ...(pivotDefinition.subtotals
          ? Array.from({ length: Math.max(0, C.length - 1) }, (T, N) =>
              C.slice(0, N + 1),
            )
          : []),
      ].map((parts) => ({
        key: JSON.stringify(parts),
        parts: parts,
        subtotal: parts.length < C.length,
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
  for (const row of rows) {
    const T = f(h(o, row)),
      N = f(h(a, row));
    (T.forEach((event) => p.set(event.key, event)),
      N.forEach((event) => m.set(event.key, event)));
    const R =
      Math.max(1, o.length) +
      (m.size + +(!!pivotDefinition.grandTotals && !!a.length)) * s.length;
    (p.size + 2) * R > 1e5 &&
      throwErrorMessage(
        "The pivot grouping would exceed the 100,000-cell output budget. Reduce group fields or filter the source.",
      );
    for (const L of T) {
      w(g, L.key, row.values);
      for (const $ of N) w(v, JSON.stringify([L.key, $.key]), row.values);
    }
    for (const L of N) w(b, L.key, row.values);
    l(y, row.values);
  }
  (m.size || m.set("[]", { key: "[]", parts: [], subtotal: false }),
    !p.size &&
      !o.length &&
      p.set("[]", { key: "[]", parts: [], subtotal: false }));
  function E(C, T) {
    if (!pivotDefinition.subtotals && !pivotDefinition.sorts?.length)
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
        const G = pivotDefinition.sorts?.find(
            (sort) => sort.field === T[U.length],
          ),
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
  const A = E(m, pivotDefinition.columnFields),
    k = E(p, pivotDefinition.rowFields),
    F = (C) =>
      C.name ??
      (C.calculated
        ? `Calculated ${C.field}`
        : `${C.operation === "distinctCount" ? "Distinct count" : C.operation[0].toUpperCase() + C.operation.slice(1)} of ${C.field}`),
    I = (C) => (isNullOrEmptyValue(C) ? "(blank)" : String(C)),
    O = pivotDefinition.rowFields.length
      ? pivotDefinition.rowFields.map((rowField) => {
          const T = pivotDefinition.dateGroups?.find(
            (dateGroup) => dateGroup.field === rowField,
          );
          return T ? `${rowField} (${T.unit})` : rowField;
        })
      : ["Row labels"];
  for (const C of A)
    for (const T of s)
      O.push(
        [...C.parts.map(I), ...(C.subtotal ? ["Subtotal"] : []), F(T)].join(
          " / ",
        ),
      );
  if (pivotDefinition.grandTotals && a.length)
    for (const C of s) O.push(`Grand Total / ${F(C)}`);
  s.forEach((C, T) => {
    if (C.showAs && C.showAs !== "value")
      for (let N = O.length - 1; N >= Math.max(1, o.length); N--)
        (N - Math.max(1, o.length)) % s.length === T &&
          (O[N] =
            `${O[N]} (% of ${C.showAs === "percentGrand" ? "grand" : C.showAs === "percentRow" ? "row" : "column"} total)`);
  });
  const rowAxes = [],
    columnAxes = [
      ...A,
      ...(pivotDefinition.grandTotals && a.length ? [null] : []),
    ],
    baseMatrix = [O],
    D = (C, T, N, R) => {
      const showAs = s[T].showAs;
      if (!showAs || showAs === "value" || C === null) return C;
      const $ = c(
        showAs === "percentGrand" ? y : showAs === "percentRow" ? N : R,
        T,
      );
      return typeof C == "number" && typeof $ == "number" && $ !== 0
        ? C / $
        : "#DIV/0!";
    },
    matrix = [O],
    subtotalRows = [];
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
    for (const columnAx of columnAxes)
      s.forEach((L, $) => {
        const U = c(
          columnAx
            ? v.get(JSON.stringify([C.key, columnAx.key]))
            : g.get(C.key),
          $,
        );
        (N.push(U),
          T.push(D(U, $, g.get(C.key), columnAx ? b.get(columnAx.key) : y)));
      });
    (C.subtotal && subtotalRows.push(matrix.length),
      matrix.push(T),
      baseMatrix.push(N),
      rowAxes.push(C));
  }
  if (pivotDefinition.grandTotals && o.length) {
    const C = ["Grand Total", ...Array(o.length - 1).fill(null)],
      T = [...C];
    for (const columnAx of columnAxes)
      s.forEach((R, L) => {
        const $ = c(columnAx ? b.get(columnAx.key) : y, L);
        (T.push($), C.push(D($, L, y, columnAx ? b.get(columnAx.key) : y)));
      });
    (matrix.push(C),
      baseMatrix.push(T),
      rowAxes.push(null),
      subtotalRows.push(matrix.length - 1));
  }
  return {
    matrix: matrix,
    baseMatrix: baseMatrix,
    rowAxes: rowAxes,
    columnAxes: columnAxes,
    chartRows: k
      .filter((C) => !C.subtotal)
      .map((C) => ({
        label: C.parts.map(I).join(" / ") || "Total",
        values: s.map((T, N) => D(c(g.get(C.key), N), N, g.get(C.key), y)),
      })),
    subtotalRows: subtotalRows,
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
      const selected = requireBoundedList(n.selected, 500).map(
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
        new Set(selected.map(buildSlicerItemKey)).size !== selected.length &&
          throwErrorMessage("Slicer selections must be unique."),
        { field: requirePivotFieldName(n.field), selected: selected }
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
    const chart = e.chart;
    if (
      !isNonArrayObject(chart) ||
      !["column", "bar", "line"].includes(String(chart.kind)) ||
      !Number.isInteger(chart.valueIndex) ||
      Number(chart.valueIndex) < 0 ||
      Number(chart.valueIndex) >= Number(e.values?.length)
    )
      return throwErrorMessage(
        "PivotChart requires a supported chart and an existing value field.",
      );
    r.chart = { kind: chart.kind, valueIndex: chart.valueIndex };
  }
  return r;
}
function validatePivotFieldReferences(e, fields) {
  const t = new Set(
      e.calculatedFields?.map((calculatedField) => calculatedField.name),
    ),
    n = [
      ...e.rowFields,
      ...e.columnFields,
      ...(e.filters ?? []).map((a) => a.field),
      ...(e.slicers ?? []).map((a) => a.field),
      ...(e.timelines ?? []).map((a) => a.field),
      ...(e.dateGroups ?? []).map((a) => a.field),
      ...(e.sorts ?? []).map((a) => a.field),
      ...e.values.filter((value) => !t.has(value.field)).map((a) => a.field),
      ...(e.calculatedFields ?? []).flatMap((a) => [
        a.field,
        ...(a.otherField ? [a.otherField] : []),
      ]),
    ];
  for (const a of n)
    fields.includes(a) ||
      throwErrorMessage(
        `Broken pivot field “${a}”: source field was not found. Restore the header or remove/reselect this field before refreshing. Historical totals are not current.`,
      );
  for (const a of e.sorts ?? [])
    [...e.rowFields, ...e.columnFields].includes(a.field) ||
      throwErrorMessage(
        `Sort field “${a.field}” must be a row or column group.`,
      );
}
function normalizeTimelineDate(e, dateSystem, field, n) {
  if (e === null || e === "") return null;
  if (dateSystem === "iso")
    return isIsoDateString(e)
      ? e
      : throwErrorMessage(
          `Timeline field “${field}” has an invalid ISO date on source row ${n + 1}.`,
        );
  if (
    typeof e != "number" ||
    !Number.isFinite(e) ||
    e < (dateSystem === "1900" ? 1 : 0) ||
    e > 2958465
  )
    return throwErrorMessage(
      `Timeline field “${field}” requires ${dateSystem} date serials on source row ${n + 1}.`,
    );
  const o = Math.floor(e);
  if (dateSystem === "1900" && o === 60) return "1900-02-29";
  const a = new Date(
    (dateSystem === "1900" ? Date.UTC(1899, 11, 31) : Date.UTC(1904, 0, 1)) +
      (o - (dateSystem === "1900" && o >= 60 ? 1 : 0)) * 864e5,
  );
  return a.getUTCFullYear() > 9999
    ? throwErrorMessage("Timeline date exceeds year 9999.")
    : a.toISOString().slice(0, 10);
}
function filterPivotSourceRows(pivotDefinition, fields, t) {
  validatePivotFieldReferences(pivotDefinition, fields);
  const n = (pivotDefinition.slicers ?? []).map((a) => ({
      index: fields.indexOf(a.field),
      keys: new Set(a.selected.map(buildSlicerItemKey)),
    })),
    o = (pivotDefinition.timelines ?? []).map((a) => ({
      ...a,
      index: fields.indexOf(a.field),
    }));
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
    const timelineDate = normalizeTimelineDate(n, r.dateSystem, r.field, 0);
    if (!timelineDate) continue;
    const a = Number(timelineDate.slice(0, 4)),
      s = Number(timelineDate.slice(5, 7)),
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
function normalizeRowProvenance(workbook, r) {
  const sheets = workbook.sheets;
  if (
    workbook.rowOrigins === undefined &&
    !sheets.some((sheet) => sheet.rowOriginIds !== undefined)
  )
    return;
  if (
    !Array.isArray(workbook.rowOrigins) ||
    workbook.rowOrigins.length > workbookLimits.maxCells
  )
    throw Error(
      "Row provenance requires a registry of at most 100,000 source rows.",
    );
  const n = new Set();
  r.rowOrigins = workbook.rowOrigins.map((rowOrigin) => {
    if (
      !isPlainDataObject(rowOrigin) ||
      !isBoundedString(rowOrigin.id, 100) ||
      n.has(rowOrigin.id) ||
      !isBoundedString(rowOrigin.sourceId, 100) ||
      !isBoundedString(rowOrigin.sourceName, 255) ||
      !isBoundedString(rowOrigin.sourceSheet, 31) ||
      !["file", "snapshot"].includes(String(rowOrigin.kind)) ||
      !Number.isInteger(rowOrigin.originalRow) ||
      Number(rowOrigin.originalRow) < 1 ||
      Number(rowOrigin.originalRow) > workbookLimits.maxRows
    )
      throw Error("Invalid, duplicated or out-of-bounds original source row.");
    return (
      n.add(rowOrigin.id),
      {
        id: rowOrigin.id,
        sourceId: rowOrigin.sourceId,
        sourceName: rowOrigin.sourceName,
        sourceSheet: rowOrigin.sourceSheet,
        originalRow: rowOrigin.originalRow,
        kind: rowOrigin.kind,
      }
    );
  });
  let o = 0;
  for (let a = 0; a < sheets.length; a++) {
    const rowOriginIds = sheets[a].rowOriginIds;
    if (rowOriginIds === undefined) continue;
    if (!isPlainDataObject(rowOriginIds))
      throw Error("Source row associations must be a plain index map.");
    const d = {};
    for (const [u, l] of Object.entries(rowOriginIds)) {
      if (
        !/^(0|[1-9]\d*)$/.test(u) ||
        Number(u) >= Number(sheets[a].rowCount) ||
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
function findRowOrigin(workbook, r, t) {
  const n = workbook.sheets.find((sheet) => sheet.id === r)?.rowOriginIds?.[t],
    o = n && workbook.rowOrigins?.find((rowOrigin) => rowOrigin.id === n);
  return o ? { ...o } : null;
}
function validateRowOrigins(workbook, r) {
  if (!workbook.rowOrigins?.length) return;
  const t = new Map((r.rowOrigins ?? []).map((o) => [o.id, o]));
  for (const rowOrigin of workbook.rowOrigins)
    if (
      !t.has(rowOrigin.id) ||
      [
        "id",
        "sourceId",
        "sourceName",
        "sourceSheet",
        "originalRow",
        "kind",
      ].some((a) => t.get(rowOrigin.id)[a] !== rowOrigin[a])
    )
      throw Error(
        "Original source row identities cannot be changed or removed. Keep the source registry and change only current row associations.",
      );
}
const basicAggregateOperations = ["sum", "count", "average", "min", "max"],
  pivotAggregateOperations = [...basicAggregateOperations, "distinctCount"];

function assertUniqueTableNames(workbook) {
  const r = new Set(),
    t = new Set();
  for (const sheet of workbook.sheets)
    for (const o of sheet.tables ?? [])
      ((r.has(o.name.toLowerCase()) || t.has(o.id)) &&
        throwWorkbookError(
          "Table names and IDs must be unique across the workbook.",
        ),
        r.add(o.name.toLowerCase()),
        t.add(o.id));
}
function assertWithinCellBudget(workbook) {
  return (
    workbook.sheets.reduce((r, t) => r + Object.keys(t.cells).length, 0) >
      workbookLimits.maxCells &&
      throwWorkbookError(
        "The result would exceed the 100,000 populated-cell workbook budget.",
      ),
    workbook
  );
}
function findSourceFieldIndex(fields, r) {
  const t = fields.indexOf(r);
  return t >= 0
    ? t
    : throwWorkbookError(
        `Source field “${r}” was not found. Refresh the field selection.`,
      );
}
function buildUniqueSheetName(workbook, r) {
  const t =
    r
      .replace(/[\\/\[\]*?:]/g, " ")
      .trim()
      .slice(0, 31) || "Analysis";
  let n = t,
    o = 2;
  for (
    ;
    workbook.sheets.some(
      (sheet) => sheet.name.toLowerCase() === n.toLowerCase(),
    );
  )
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
function findPivotDefinition(workbook, pivotId) {
  const sheet = workbook.sheets.find((sheet) =>
      sheet.pivots?.some((pivot) => pivot.id === pivotId),
    ),
    definition = sheet?.pivots?.find((pivot) => pivot.id === pivotId);
  return (
    (!sheet || !definition) &&
      throwWorkbookError("Pivot definition not found."),
    { owner: sheet, definition: definition }
  );
}
function buildPivotStatusRecord(workbook, definition, pivotRefreshState) {
  return {
    workbookId: workbook.id,
    revision: workbook.revision,
    pivotId: definition.id,
    name: definition.name,
    source: structuredClone(definition.source),
    sourceRevision: definition.sourceRevision,
    refreshedAt: definition.refresh?.at ?? null,
    sourceFingerprint: definition.refresh?.sourceFingerprint ?? null,
    status: pivotRefreshState.state,
    diagnostic: pivotRefreshState.reason,
    outputRange: definition.outputRange,
    rowFields: [...definition.rowFields],
    columnFields: [...definition.columnFields],
    values: structuredClone(definition.values),
    filters: structuredClone(definition.filters ?? []),
    slicers: structuredClone(definition.slicers ?? []),
    timelines: structuredClone(definition.timelines ?? []),
    sorts: structuredClone(definition.sorts ?? []),
    dateGroups: structuredClone(definition.dateGroups ?? []),
    calculatedFields: structuredClone(definition.calculatedFields ?? []),
    chart: structuredClone(definition.chart ?? null),
    population: definition.refresh
      ? {
          source: definition.refresh.sourceRows,
          included: definition.refresh.includedRows,
          excluded: definition.refresh.excludedRows,
        }
      : null,
    measures: structuredClone(definition.refresh?.measures ?? []),
    semantics:
      "Controls intersect. Slicer items use type-sensitive, case-insensitive equality; an empty selection includes no records. Hidden source rows are included. Percentages use filtered aggregate totals. Drill-through returns numerator records, including blank/nonnumeric measure inputs; it does not expand denominator populations.",
  };
}
const tableMutationLimits = {
  changedCells: 500,
  appendedRows: 100,
  formulaLength: 4096,
};
function validateCalculatedColumnFormula(refersTo, r = false) {
  (typeof refersTo != "string" ||
    refersTo.length > tableMutationLimits.formulaLength ||
    refersTo !== refersTo.trim() ||
    !refersTo.startsWith("=")) &&
    throwWorkbookError(
      "Calculated column formulas must begin with = and contain at most 4,096 characters, without outer whitespace.",
    );
  try {
    validateNameFormulaText(refersTo);
  } catch (error) {
    throwWorkbookError(
      `Calculated column formula: ${error.message.replace(/^XLSX defined names: /, "").replace(/ Import\/export stopped.*$/, "")}`,
    );
  }
  return (
    !r &&
      /#REF!/.test(
        stripFormulaStringLiterals(refersTo.replace(/"(?:[^"]|"")*"/g, '""')),
      ) &&
      throwWorkbookError(
        "Repair broken #REF! references before saving a calculated column rule.",
      ),
    refersTo
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
      let col = 0;
      for (const b of m[1].toUpperCase()) col = col * 26 + b.charCodeAt(0) - 64;
      const row = Number(m[2]);
      if (col > 16384 || row > 1048576)
        throw Error("External reference exceeds worksheet bounds.");
      return { row: row, col: col };
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
function tokenizeExternalReferences(raw) {
  const r = [];
  for (let end = 0; end < raw.length;) {
    if (raw[end] === '"') {
      for (end++; end < raw.length;)
        if (raw[end++] === '"')
          if (raw[end] === '"') end++;
          else break;
      continue;
    }
    const start = end,
      o = /^'((?:[^']|'')*)'\s*!/.exec(raw.slice(end)),
      a = /^\[([^\[\]\r\n]+)\]([A-Za-z_][A-Za-z0-9_.]*)!/.exec(raw.slice(end));
    let s,
      d = 0;
    if (
      (o
        ? ((s = o[1].replace(/''/g, "'")), (d = o[0].length))
        : a && ((s = `[${a[1]}]${a[2]}`), (d = a[0].length)),
      s)
    ) {
      const u = /^(.*\[[^\[\]\r\n]+\])([^\[\]:\r\n]+)$/.exec(s),
        l = a1ReferenceAnchorPattern.exec(raw.slice(end + d));
      if (u && l && (!end || !/[A-Za-z0-9_.\]]/.test(raw[end - 1])))
        try {
          (parseExternalLinkRange(l[1]),
            (end += d + l[0].length),
            r.push({
              start: start,
              end: end,
              token: raw.slice(start, end),
              target: u[1],
              sheet: u[2],
              range: l[1].toUpperCase(),
            }));
          continue;
        } catch {}
      end += d;
      continue;
    }
    if (raw[end] === "[") {
      let u = 0;
      do (raw[end] === "[" && u++, raw[end] === "]" && u--, end++);
      while (end < raw.length && u > 0);
    } else if (raw[end] === "'") {
      for (end++; end < raw.length;)
        if (raw[end++] === "'")
          if (raw[end] === "'") end++;
          else break;
    } else end++;
  }
  return r;
}
function rewriteExternalReferences(raw, r) {
  const t = tokenizeExternalReferences(raw);
  for (const n of t.reverse()) {
    const o = r(n);
    o !== undefined && (raw = raw.slice(0, n.start) + o + raw.slice(n.end));
  }
  return raw;
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
function attachWorkbookLinkMetadata(workbook, r) {
  if (workbook.workbookLinks === undefined) return;
  if (
    workbook.workbookLinks &&
    typeof workbook.workbookLinks == "object" &&
    validatedWorkbookLinkRegistry.has(workbook.workbookLinks)
  ) {
    r.workbookLinks = workbook.workbookLinks;
    return;
  }
  if (
    JSON.stringify(workbook.workbookLinks).length >
    workbookLinkLimits.metadataCharacters
  )
    throw Error("Workbook link metadata exceeds 2 million characters.");
  const t = workbookLinkPackageSchema.parse(workbook.workbookLinks);
  if (
    new Set(t.sources.map((source) => source.target)).size !== t.sources.length
  )
    throw Error("Workbook link targets must be unique.");
  if (
    new Set(t.packageTargets?.map((packageTarget) => packageTarget.key))
      .size !== (t.packageTargets?.length ?? 0)
  )
    throw Error("Package link identities must be unique.");
  let n = 0,
    o = 0;
  for (const source of t.sources) {
    if (
      new Set(source.sheets.map((sheet) => sheet.name.toLowerCase())).size !==
      source.sheets.length
    )
      throw Error("Linked worksheet names must be unique.");
    for (const sheet of source.sheets) {
      o++;
      for (const cell of Object.keys(sheet.cells))
        (parseExternalLinkRange(cell), n++);
    }
  }
  if (n > workbookLinkLimits.cells || o > workbookLinkLimits.sheets)
    throw Error("Workbook link cache exceeds its cell or sheet limit.");
  ((r.workbookLinks = deepFreezeValue(t)),
    validatedWorkbookLinkRegistry.add(r.workbookLinks));
}
function encodeWorkbookFingerprint(workbook) {
  const r = new TextEncoder().encode(
    JSON.stringify([
      workbook.id,
      workbook.name,
      workbook.revision,
      workbook.calculation,
      workbook.datePolicy,
      workbook.definedNames,
      workbook.workbookLinks,
      workbook.sheets.map((sheet) => [
        sheet.id,
        sheet.name,
        sheet.rowCount,
        sheet.colCount,
        sheet.hidden,
        sheet.veryHidden,
        sheet.hiddenRows,
        sheet.hiddenCols,
        sheet.tables,
        sheet.merges,
        sheet.outlineGroups,
        Object.entries(sheet.cells).map(([n, cell]) => [
          n,
          cell.raw,
          !!cell.literal,
        ]),
      ]),
    ]),
  );
  return sha256HexDigest(r);
}
function buildExternalLinkAliases(workbookInput) {
  const sheets = [],
    values = new Map(),
    n = new Map(),
    o = new Set(workbookInput.sheets.map((sheet) => sheet.name.toLowerCase()));
  for (const [d, workbook] of (
    workbookInput.workbookLinks?.sources ?? []
  ).entries())
    for (const [l, c] of workbook.sheets.entries()) {
      let h = `_LG_LINK_${d}_${l}`;
      for (; o.has(h.toLowerCase());) h += "_";
      o.add(h.toLowerCase());
      let f = `linked:${d}:${l}`;
      for (; workbookInput.sheets.some((sheet) => sheet.id === f);) f += "_";
      const cells = {};
      for (const [m, v] of Object.entries(c.cells))
        ((cells[m] = { raw: v.value, literal: true }),
          values.set(`${f}:${m}`, v));
      (sheets.push({
        id: f,
        name: h,
        rowCount: 1048576,
        colCount: 16384,
        cells: cells,
      }),
        n.set(JSON.stringify([workbook.target, c.name.toLowerCase()]), h));
    }
  return {
    sheets: sheets,
    values: values,
    rewrite: (raw) =>
      rewriteExternalReferences(raw, (u) => {
        const workbook = workbookInput.workbookLinks?.sources.find(
            (source) => source.target === u.target,
          ),
          c = workbook?.sheets.find(
            (sheet) => sheet.name.toLowerCase() === u.sheet.toLowerCase(),
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
        const char = $[Q];
        if (char === "~") {
          const J = $[Q + 1];
          J && ["*", "?", "~"].includes(J)
            ? (U.push({ kind: "literal", char: J }), Q++, (G = true))
            : U.push({ kind: "literal", char: char });
        } else
          char === "*" || char === "?"
            ? ((G = true),
              (char === "?" || U.at(-1)?.kind !== "many") &&
                U.push({ kind: char === "*" ? "many" : "one" }))
            : U.push({ kind: "literal", char: char });
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
  const b = s
    .slice(1)
    .map((T, N) =>
      T.map((R, L) =>
        l[L] < 0 ? t.criterionFormula?.(r[2], N + 1, L) : undefined,
      ),
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
  isFormulaCellValue = (cell) =>
    typeof cell?.raw == "string" && cell.raw.startsWith("=") && !cell.literal;
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
function findScopedDefinedName(e, workbook, t) {
  if (
    e.cell ||
    (e.call &&
      !e.qualifier &&
      builtinFunctionNames.includes(normalizeFunctionName(e.name)))
  )
    return;
  const n = e.qualifier
    ? workbook.sheets.find(
        (sheet) => sheet.name.toLowerCase() === e.qualifier.toLowerCase(),
      )?.id
    : t;
  if (n)
    return (
      workbook.definedNames?.find(
        (definedName) =>
          definedName.scopeSheetId === n &&
          definedName.name.toLowerCase() === e.name.toLowerCase(),
      ) ??
      workbook.definedNames?.find(
        (definedName) =>
          !definedName.scopeSheetId &&
          definedName.name.toLowerCase() === e.name.toLowerCase(),
      )
    );
}
function resolveDefinedNameValue(e, workbook, t, n = [], o) {
  if (n.length > 32) throw new Error("#LIMIT!");
  const a = replaceTextSpans(e, m5(e), (s) => {
    if (s.call) return s.text;
    const scopedDefinedName = findScopedDefinedName(s, workbook, t);
    if (!scopedDefinedName || parseFormulaTree(scopedDefinedName.refersTo))
      return s.text;
    const u = definedNameScopeKey(scopedDefinedName);
    if (n.includes(u)) throw new Error("#CYCLE!");
    const l =
      scopedDefinedName.scopeSheetId ??
      (s.qualifier
        ? workbook.sheets.find(
            (sheet) => sheet.name.toLowerCase() === s.qualifier.toLowerCase(),
          ).id
        : t);
    if (o && !o(l)) throw new Error("#BLOCKED!");
    let definedNameValue = resolveDefinedNameValue(
      scopedDefinedName.refersTo.slice(1),
      workbook,
      l,
      [...n, u],
      o,
    );
    if (l !== t) {
      const sheet = workbook.sheets.find((sheet) => sheet.id === l),
        f = m5(definedNameValue);
      definedNameValue = replaceTextSpans(definedNameValue, f, (p) => {
        const m = f.indexOf(p),
          v = f[m - 1],
          g = v?.cell && definedNameValue.slice(v.end, p.start).trim() === ":";
        return p.cell && !p.qualifier && !g
          ? formatQuotedSheetPrefix(sheet.name) + p.text
          : p.text;
      });
    }
    return `(${definedNameValue})`;
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

/* Also used by the workers. */
export {
  FormulaErrorClass,
  a1CellPattern,
  a1ReferenceAnchorPattern,
  aggregateFunctionNames,
  databaseFunctionLimits,
  linkedCellValueSchema,
  tU,
  workbookLinkPackageSchema,
  workbookLinkSourceSchema,
};

/* Also used by the workers. */
export {
  assertPlainTextLabel,
  isBoundedString,
  isIsoDateString,
  requireBoundedList,
  requirePivotFieldName,
};
