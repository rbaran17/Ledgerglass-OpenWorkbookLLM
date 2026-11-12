import {
  transformationRecipeLimits,
  fingerprintCanonicalValue,
  normalizeDatePolicy,
  dateSystemMaxSerial,
  parseDateText,
} from "./01-core-schemas-and-hashing.js";
import { workbookLimits } from "./02-financial-functions.js";
import {
  textJoinSpecSchema,
  validateFieldTransformRecipe,
  createWorksheetRecord,
  findRowOrigin,
  calculateWorkbookWithSettings,
  validateFieldTransformEvidence,
  textInferenceLimits,
  sourceRowValueSchema,
  _t,
  parseRangeAddress,
  normalizeCellRecordInput,
  normalizeWorkbookInput,
  columnIndexToLetters,
  checkDataValidationBatch,
} from "./03-workbook-model.js";
import {
  readSourceHeaders,
  matchKeyPopulations,
  inspectFormulaDependencies,
} from "./12-formula-dependency-inspector.js";
import {
  validateTextToColumnsRequest,
  assertSelectionInAuthorizedSheet,
  splitDelimitedCell,
  throwValidationError,
  parseImportedCellValue,
  rangesOverlap,
  dataModelBudgets,
  formatCellTextValue,
  lg,
  validateTextToColumnsPermissions,
  assertOutputNameUnreferenced,
} from "./15-data-model-dialog.js";

function discoverArrayFormulaAnchors(e, r, t) {
  const n = [];
  let o = 0,
    a = 0;
  for (const h of new Map(r.map((f) => [f.sheet.id, f.sheet])).values())
    for (const [f, p] of Object.entries(h.cells)) {
      if (typeof p.raw != "string" || !p.raw.startsWith("=") || p.literal)
        continue;
      (++o > dataModelBudgets.formulaCells ||
        (a += p.raw.length) > dataModelBudgets.formulaCharacters) &&
        throwValidationError(
          "Array checking supports at most 1,000 formulas and 100,000 formula characters on the participating worksheets.",
        );
      const m = _t(f);
      r.some((g) => {
        if (g.sheet.id !== h.id) return false;
        const b = parseRangeAddress(g.range),
          y = Math.max(m.row, b.firstRow),
          w = Math.max(m.col, b.firstCol);
        return (
          y <= b.lastRow &&
          w <= b.lastCol &&
          (y - m.row + 1) * (w - m.col + 1) <= 1e4
        );
      }) && n.push({ sheet: h, anchor: f });
    }
  if (
    (n.length > dataModelBudgets.discoveryAnchors &&
      throwValidationError(
        "The selection has more than 50 possible array anchors. Choose a new output sheet or a smaller reviewed region.",
      ),
    !n.length)
  )
    return;
  const s = new Set(),
    d = (h, f) => JSON.stringify([h, f]),
    u = {
      ...e,
      sheets: e.sheets.map((h) => (h.veryHidden ? { ...h, hidden: true } : h)),
    };
  for (const h of n) {
    assertSelectionInAuthorizedSheet(h.sheet, h.anchor, t);
    const f = inspectFormulaDependencies(u, {
        sheetId: h.sheet.id,
        address: h.anchor,
        direction: "precedents",
        depth: 8,
        includeHidden: t.includeHidden,
        allowedSheetIds: t.allowedSheetIds,
        limits: { maxFormulaCells: 1e3, maxNodes: 500, maxReferences: 2e3 },
      }),
      p = [f.target, ...f.nodes],
      m = new Set(p.map((v) => v.key));
    (f.coverage.truncated ||
      f.coverage.depthLimited ||
      f.cycles.length ||
      f.issues.some((v) => m.has(v.from))) &&
      throwValidationError(
        "Array discovery requires a complete, authorized static dependency closure. Resolve hidden, dynamic, cyclic or unsupported formula dependencies before splitting.",
      );
    for (const v of p) {
      const g = e.sheets.find((b) => b.id === v.sheetId);
      (assertSelectionInAuthorizedSheet(g, v.address, t),
        (!g.cells[v.address] || g.cells[v.address].raw === null) &&
          throwValidationError(
            "Array discovery cannot infer blank or possible spill-child dependencies. Use a reviewed values snapshot before splitting.",
          ),
        s.add(d(v.sheetId, v.address)),
        s.size > dataModelBudgets.dependencyCells &&
          throwValidationError(
            "Array discovery exceeds 2,000 dependency cells. Use a smaller reviewed values snapshot.",
          ));
    }
  }
  const l = {
      ...e,
      sheets: e.sheets
        .filter((h) => t.allowedSheetIds.includes(h.id))
        .map((h) => ({
          ...h,
          cells: Object.fromEntries(
            Object.entries(h.cells).filter(([f]) => s.has(d(h.id, f))),
          ),
        })),
    },
    c = calculateWorkbookWithSettings(l);
  for (const { sheet, anchor } of n) {
    const p = c.getSpillRange(sheet.id, anchor);
    p &&
      r.some((m) => m.sheet.id === sheet.id && rangesOverlap(m.range, p)) &&
      throwValidationError(
        "The selected source or destination intersects a generated array. Convert its results to ordinary values explicitly before splitting.",
      );
  }
  c.getOperationalErrors().length &&
    throwValidationError(
      "Array discovery encountered an operational calculation error. Resolve the uncertain formula state before splitting text.",
    );
}
function buildTextSplitResult(e, r, t = {}) {
  var w, A, k, F, I;
  const n = validateTextToColumnsPermissions(e, t),
    o = validateTextToColumnsRequest(r),
    a = e.sheets.find((O) => O.id === o.sourceSheetId);
  (a || throwValidationError("Choose an existing source worksheet."),
    assertSelectionInAuthorizedSheet(a, o.sourceRange, n));
  const s = parseRangeAddress(o.sourceRange),
    d = s.lastRow - s.firstRow + 1;
  ((s.firstCol !== s.lastCol || d > dataModelBudgets.sourceCells) &&
    throwValidationError(
      "Select one source column containing at most 500 cells.",
    ),
    (w = a.merges) != null &&
      w.some((O) => rangesOverlap(O, o.sourceRange)) &&
      throwValidationError("Unmerge the source before splitting text."));
  const u = [],
    l = [],
    c = [];
  let h = 0,
    f = 0;
  for (let O = s.firstRow; O <= s.lastRow; O++) {
    const j = lg(O, s.firstCol),
      z = a.cells[j];
    z &&
      typeof z.raw == "string" &&
      z.raw.startsWith("=") &&
      !z.literal &&
      throwValidationError(
        `Source ${j} contains a formula. Copy its results to ordinary values explicitly before splitting.`,
      );
    const V = formatCellTextValue(z?.raw ?? null);
    (f += V.length) > dataModelBudgets.sourceCharacters &&
      throwValidationError(
        "Selected source text exceeds the 200,000-character processing budget.",
      );
    let D;
    try {
      D = splitDelimitedCell(V, o.delimiter);
    } catch (B) {
      throwValidationError(`Source ${j}: ${B.message}`);
    }
    ((h = Math.max(h, D.length)),
      l.push(D),
      u.push({
        address: j,
        ...(z ? { cell: structuredClone(z) } : {}),
        text: V,
        fields: D.length,
      }));
  }
  (h * d > dataModelBudgets.outputCells &&
    throwValidationError(
      "Split output may contain at most 500 cells. Select fewer source rows or fields.",
    ),
    (o.formats?.length ?? 0) > h &&
      throwValidationError(
        "A column conversion refers to a field outside the split result.",
      ));
  const p = Array.from({ length: h }, (O, j) => {
    return o.formats?.[j] ?? "text";
  });
  let m,
    v,
    g = { row: 0, col: 0 };
  if (o.destination.kind === "new-sheet")
    ((!n.allowNewSheet || ((A = e.structureProtection) != null && A.enabled)) &&
      throwValidationError(
        "Creating an output worksheet is not permitted while workbook structure is protected or the host disallows new worksheets.",
      ),
      (v = o.destination.name),
      e.sheets.some((O) => O.name.toLowerCase() === v.toLowerCase()) &&
        throwValidationError(
          "The output worksheet name is already used. Choose another name.",
        ),
      assertOutputNameUnreferenced(e, v));
  else {
    const O = o.destination;
    ((m = e.sheets.find((j) => j.id === O.sheetId)),
      m || throwValidationError("Choose an existing output worksheet."),
      (g = _t(o.destination.address)),
      (v = m.name),
      (k = m.protection) != null &&
        k.enabled &&
        throwValidationError("The output worksheet is protected."));
  }
  (g.row + d > workbookLimits.maxRows || g.col + h > workbookLimits.maxCols) &&
    throwValidationError("Split output exceeds the worksheet boundary.");
  const b = `${lg(g.row, g.col)}:${lg(g.row + d - 1, g.col + h - 1)}`;
  (m &&
    (assertSelectionInAuthorizedSheet(m, b, n),
    m.id === a.id &&
      rangesOverlap(o.sourceRange, b) &&
      throwValidationError(
        "Choose a separate destination so the original source cells remain intact.",
      ),
    (((F = m.tables) != null && F.some((O) => rangesOverlap(O.range, b))) ||
      ((I = m.merges) != null && I.some((O) => rangesOverlap(O, b))) ||
      [...(m.pivots ?? []), ...(m.queries ?? [])].some((O) =>
        rangesOverlap(O.outputRange, b),
      )) &&
      throwValidationError(
        "Choose an output outside tables, merged areas and generated analysis results.",
      ),
    Object.keys(m.cells).some((O) => rangesOverlap(O, b)) &&
      throwValidationError(
        "The destination contains cells or their formatting. Choose an empty range or new worksheet.",
      ),
    [...(e.reviewThreads ?? []), ...(e.formulaWatches ?? [])].some(
      (O) =>
        !O.orphaned &&
        O.anchor.sheetId === m.id &&
        rangesOverlap(O.anchor.address, b),
    ) &&
      throwValidationError(
        "The destination contains a saved watch or discussion anchor. Choose another range to preserve its evidence context.",
      )),
    discoverArrayFormulaAnchors(
      e,
      [
        { sheet: a, range: o.sourceRange },
        ...(m ? [{ sheet: m, range: b }] : []),
      ],
      n,
    ));
  const y = [];
  for (let O = 0; O < d; O++) {
    l[O].length !== h &&
      c.push({
        code: "padded-row",
        sourceAddress: u[O].address,
        message: `This record has ${l[O].length} fields; ${h - l[O].length} missing fields are padded with empty text.`,
      });
    for (let j = 0; j < h; j++) {
      const z = l[O][j] ?? "",
        V = o.header === "first-row" && O === 0 ? "text" : p[j],
        D = parseImportedCellValue(z, V);
      D.precision &&
        c.push({
          code: "numeric-text-retained",
          sourceAddress: u[O].address,
          message: `Field ${j + 1} remains text because numeric conversion would exceed the supported finite precision.`,
        });
      const B = `Text to Columns derivation
Workbook ID: ${JSON.stringify(e.id)}
Source revision: ${e.revision}
Worksheet ID: ${JSON.stringify(a.id)}
Worksheet name: ${JSON.stringify(a.name)}
Source cell: ${u[O].address}
Field: ${j + 1}
Conversion: ${V}
${j >= l[O].length ? "Padded missing field." : "Original source cell and evidence remain in place."}`,
        _ = normalizeCellRecordInput({
          raw: D.raw,
          note: B,
          ...(typeof D.raw == "string"
            ? {
                numFmt: "@",
                ...(D.raw.startsWith("=") ? { literal: true } : {}),
              }
            : {}),
          ...(o.header === "first-row" && O === 0
            ? { style: { bold: true } }
            : {}),
        });
      y.push({
        sourceAddress: u[O].address,
        address: lg(g.row + O, g.col + j),
        field: z,
        padded: j >= l[O].length,
        format: V,
        type:
          typeof D.raw == "number"
            ? "number"
            : typeof D.raw == "boolean"
              ? "boolean"
              : "text",
        cell: _,
      });
    }
  }
  if (
    (e.sheets.reduce((O, j) => O + Object.keys(j.cells).length, 0) + y.length >
      workbookLimits.maxCells &&
      throwValidationError(
        "The split would exceed the workbook's 100,000 stored-cell budget.",
      ),
    m)
  ) {
    const O = structuredClone(e),
      j = O.sheets.find((z) => z.id === m.id);
    for (const z of y) j.cells[z.address] = z.cell;
    checkDataValidationBatch(
      O,
      y.map((z) => ({ sheetId: j.id, address: z.address, after: z.cell })),
    );
  }
  return {
    snapshot: JSON.stringify(e),
    request: o,
    context: n,
    sourceSheetName: a.name,
    targetSheetName: v,
    targetRange: b,
    width: h,
    height: d,
    formats: p,
    sources: u,
    outputs: y,
    diagnostics: c,
  };
}
function applyTextSplitPreview(e, r, t = {}) {
  r.snapshot !== JSON.stringify(e) &&
    throwValidationError(
      "The workbook changed after the split preview. Preview again.",
    );
  const n = buildTextSplitResult(e, r.request, t);
  JSON.stringify(n) !== JSON.stringify(r) &&
    throwValidationError(
      "The split preview or authorized context changed. Preview again.",
    );
  const o = structuredClone(e),
    a = r.request.destination,
    s =
      a.kind === "new-sheet"
        ? createWorksheetRecord(a.name)
        : o.sheets.find((d) => d.id === a.sheetId);
  a.kind === "new-sheet" && o.sheets.push(s);
  for (const d of r.outputs) s.cells[d.address] = structuredClone(d.cell);
  return (normalizeWorkbookInput(o), o);
}
const coerceCellValueToText = (e) =>
  e === null ? "" : typeof e == "boolean" ? (e ? "TRUE" : "FALSE") : String(e);
function applyTextCasing(e, r) {
  return r === "upper"
    ? e.toLocaleUpperCase("en-US")
    : r === "lower"
      ? e.toLocaleLowerCase("en-US")
      : r === "proper"
        ? e
            .toLocaleLowerCase("en-US")
            .replace(new RegExp("\\p{L}[\\p{L}\\p{M}]*", "gu"), (t) => {
              const [n, ...o] = [...t];
              return n.toLocaleUpperCase("en-US") + o.join("");
            })
        : e;
}
function extractInferredFieldPart(e, r) {
  if (r.column >= e.length)
    throw Error("The inferred rule refers to a missing field.");
  const t = coerceCellValueToText(e[r.column]).trim();
  if (r.mode === "whole") return t;
  if (r.mode === "slice") {
    const s = [...t];
    if (s.length < r.start + r.length)
      throw Error(
        "The value is too short for the inferred character extraction.",
      );
    return s.slice(r.start, r.start + r.length).join("");
  }
  const n = t.indexOf(r.delimiter);
  if (r.mode === "before" || r.mode === "after") {
    if (n < 0) throw Error("The inferred separator is absent.");
    return (
      r.mode === "before" ? t.slice(0, n) : t.slice(n + r.delimiter.length)
    ).trim();
  }
  const o = r.delimiter === " " ? t.split(/ +/) : t.split(r.delimiter),
    a = r.index < 0 ? o.length + r.index : r.index;
  if (a < 0 || a >= o.length || !o[a])
    throw Error("The inferred token is absent.");
  return o[a].trim();
}
function applyInferredTextPattern(e, r) {
  const t = textJoinSpecSchema.parse(r),
    n =
      t.prefix +
      applyTextCasing(
        t.parts.map((o) => extractInferredFieldPart(e, o)).join(t.separator),
        t.casing,
      ) +
      t.suffix;
  if (n.length > 32767)
    throw Error("The inferred result exceeds the cell text limit.");
  return n;
}
function describeInferredTextRule(e) {
  const r = e.parts
    .map(
      (t) =>
        `field ${t.column + 1}${t.mode === "whole" ? "" : t.mode === "token" ? ` token ${t.index < 0 ? "last" : t.index + 1} by ${JSON.stringify(t.delimiter)}` : t.mode === "slice" ? ` characters ${t.start + 1}–${t.start + t.length}` : ` ${t.mode} ${JSON.stringify(t.delimiter)}`}`,
    )
    .join(` + ${JSON.stringify(e.separator)} + `);
  return `${e.prefix ? JSON.stringify(e.prefix) + " + " : ""}${e.casing === "none" ? "" : e.casing + "("}${r}${e.casing === "none" ? "" : ")"}${e.suffix ? " + " + JSON.stringify(e.suffix) : ""}`;
}
function inferTextPatternFromExamples(e, r, t) {
  if (!Array.isArray(t) || t.length < 2 || t.length > 10)
    throw Error("Provide 2–10 examples from distinct source rows.");
  const n = t.map((f) => sourceRowValueSchema.parse(f));
  if (new Set(n.map((f) => f.row)).size !== n.length)
    throw Error("Provide examples from distinct source rows.");
  const o = n.map((f) => ({ example: f, row: e.find((p) => p.row === f.row) }));
  if (o.some((f) => !f.row))
    throw Error("An example row is outside the selected data rows.");
  if (
    !r.length ||
    r.length > 8 ||
    new Set(r).size !== r.length ||
    r.some(
      (f) =>
        !Number.isInteger(f) ||
        f < 0 ||
        o.some((p) => f >= p.row.values.length),
    )
  )
    throw Error("Choose 1–8 available fields for inference.");
  const a = [],
    s = (f) => ({
      column: f,
      mode: "whole",
      delimiter: " ",
      index: 0,
      start: 0,
      length: 1,
    });
  for (const f of r) {
    a.push(s(f));
    for (const p of [" ", ",", "-", "/", "_", "@"]) {
      for (const m of [0, 1, -1])
        a.push({ ...s(f), mode: "token", delimiter: p, index: m });
      for (const m of ["before", "after"])
        a.push({ ...s(f), mode: m, delimiter: p });
    }
    for (const p of [1, 2, 3, 4, 5, 6, 8, 10])
      a.push({ ...s(f), mode: "slice", length: p });
  }
  const d = a.map((f) => ({ parts: [f], separator: "" }));
  for (const f of r)
    for (const p of r)
      if (f !== p)
        for (const m of [" ", "", ", ", ".", "-", "_", "/"])
          d.push({ parts: [s(f), s(p)], separator: m });
  const u = [],
    l = new Set();
  let c = 0,
    h = false;
  e: for (const f of d)
    for (const p of ["none", "upper", "lower", "proper"]) {
      if (++c > textInferenceLimits.inferencePatterns) {
        (c--, (h = true));
        break e;
      }
      const m = { ...f, casing: p, prefix: "", suffix: "" };
      let v;
      try {
        v = o.map((E) => applyInferredTextPattern(E.row.values, m));
      } catch {
        continue;
      }
      if (v.some((E) => !E) || new Set(v).size < 2) continue;
      const g = o[0].example.value,
        b = v[0];
      let y = g.indexOf(b),
        w = 0;
      for (; y >= 0 && w++ < 3;) {
        const E = {
          ...m,
          prefix: g.slice(0, y),
          suffix: g.slice(y + b.length),
        };
        if (
          E.prefix.length <= 1e3 &&
          E.suffix.length <= 1e3 &&
          o.every((A, k) => E.prefix + v[k] + E.suffix === A.example.value)
        ) {
          const A = JSON.stringify(E);
          if (
            !l.has(A) &&
            (l.add(A),
            u.push({
              id: `pattern-${u.length + 1}`,
              pattern: E,
              description: describeInferredTextRule(E),
              matchedExamples: n.length,
            }),
            u.length >= textInferenceLimits.inferenceCandidates)
          ) {
            h = true;
            break e;
          }
        }
        y = g.indexOf(b, y + 1);
      }
    }
  return { candidates: u, searched: c, truncated: h };
}
const Gh = (e, r) => `${columnIndexToLetters(r)}${e + 1}`,
  isBlankCellValue = (e) => e === null || e === "",
  inferCellValueType = (e) =>
    e === null
      ? "blank"
      : typeof e == "number"
        ? "number"
        : typeof e == "boolean"
          ? "boolean"
          : "text",
  uniqueSourceReferences = (e) => [...new Set(e.flatMap((r) => r.sources))],
  createSourcedValue = (e, r, t = inferCellValueType(e)) => ({
    value: e,
    type: t,
    sources: r,
    exception: false,
  });
function assertRecipeFieldInRange(e, r) {
  if (e.some((t) => t >= r))
    throw Error(
      `A recipe field exceeds the current ${r} columns. Review steps in order.`,
    );
}
function splitTextWithQuoteRules(e, r) {
  if (r.mode === "fixed") {
    const h = [...e],
      f = [0, ...r.breaks, h.length];
    return f
      .slice(0, -1)
      .map((p, m) => h.slice(p, Math.max(p, f[m + 1])).join(""));
  }
  const t = new Set([...r.delimiters]),
    n = [...e],
    o = [];
  let a = "",
    s = false,
    d = false,
    u = false,
    l = false;
  const c = () => {
    if ((o.push(a), (a = ""), (d = false), (u = false), o.length >= 50))
      throw Error("More than 50 split fields.");
  };
  for (let h = 0; h < n.length; h++) {
    const f = n[h];
    if (s)
      (f !== '"'
        ? (a += f)
        : n[h + 1] === '"'
          ? ((a += '"'), h++)
          : ((s = false), (d = true)),
        (l = false));
    else if (t.has(f)) ((r.collapse && l) || c(), (l = true));
    else {
      if (d) throw Error("Unexpected text after a closing quote.");
      if (f === '"') {
        if (u) throw Error("A quote must begin a field.");
        ((s = true), (u = true));
      } else ((a += f), (u = true));
      l = false;
    }
  }
  if (s) throw Error("The quoted field is not closed.");
  return (o.push(a), o);
}
function convertFieldToTypedValue(e, r, t = "1900") {
  if (r.to === "text") return coerceCellValueToText(e);
  if (isBlankCellValue(e)) return null;
  if (r.to === "boolean") {
    if (typeof e == "boolean") return e;
    if (typeof e == "string" && /^(true|false)$/i.test(e.trim()))
      return e.trim().toUpperCase() === "TRUE";
    throw Error("Expected TRUE or FALSE.");
  }
  if (r.to === "date") {
    if (
      typeof e == "number" &&
      Number.isInteger(e) &&
      e >= (t === "1904" ? 0 : 1) &&
      e <= dateSystemMaxSerial(t)
    )
      return e;
    if (
      typeof e != "string" ||
      !/^\s*(?:\d{4}-\d{1,2}-\d{1,2}|\d{1,2}\/\d{1,2}\/\d{4})\s*$/.test(e)
    )
      throw Error(
        "Expected a full-year ISO date or an explicitly ordered slash date.",
      );
    const s = parseDateText(e, { textDateOrder: r.dateOrder, dateSystem: t });
    if (s === undefined)
      throw Error(`The date is impossible or outside ${t}–9999.`);
    return s;
  }
  if (typeof e == "boolean")
    throw Error("Boolean values are not numeric fields.");
  let n;
  if (typeof e == "number") n = e;
  else {
    let d = e.trim();
    /^\(.*\)$/.test(d) && (d = "-" + d.slice(1, -1));
    const u = /[eE][+-]?\d+$/.exec(d)?.[0] ?? "",
      c = (u ? d.slice(0, -u.length) : d).split(r.decimal);
    if (c.length > 2) throw Error("More than one decimal separator.");
    const h = /^[+-]/.exec(c[0])?.[0] ?? "",
      f = c[0].slice(h.length),
      p = c[1];
    if (r.group && f.includes(r.group)) {
      const b = f.split(r.group);
      if (!/^\d{1,3}$/.test(b[0]) || b.slice(1).some((y) => !/^\d{3}$/.test(y)))
        throw Error(
          "Grouping separators must separate groups of three digits.",
        );
    }
    const m = r.group ? f.split(r.group).join("") : f;
    if (!/^\d*$/.test(m) || (p !== undefined && !/^\d+$/.test(p)) || (!m && !p))
      throw Error("The field does not match the explicit numeric separators.");
    if ((m + (p ?? "")).replace(/^0+/, "").replace(/0+$/, "").length > 15)
      throw Error(
        "More than 15 significant digits; original text is retained.",
      );
    const g = h + (m || "0") + (p === undefined ? "" : "." + p) + u;
    if (
      ((n = Number(g)),
      !Number.isFinite(n) ||
        (n === 0 && /[1-9]/.test(m + (p ?? ""))) ||
        (Number.isInteger(n) && !Number.isSafeInteger(n)))
    )
      throw Error("Numeric conversion exceeds supported precision or range.");
  }
  if (!Number.isFinite(n)) throw Error("Expected a finite number.");
  if (r.to === "integer" && !Number.isSafeInteger(n))
    throw Error("Expected a safe whole number.");
  return n;
}
function prepareTransformationInput(e, r, t) {
  var p, m;
  const n = validateFieldTransformRecipe(r),
    o = validateTextToColumnsPermissions(e, t),
    a = e.sheets.find((v) => v.id === n.sourceSheetId);
  if (!a) throw Error("Choose an existing source worksheet.");
  if (!o.allowNewSheet || ((p = e.structureProtection) != null && p.enabled))
    throw Error("Creating a derived worksheet is not permitted.");
  if (e.sheets.some((v) => v.name.toLowerCase() === n.outputName.toLowerCase()))
    throw Error("The output worksheet name is already used.");
  const s = parseRangeAddress(n.sourceRange),
    d = s.lastCol - s.firstCol + 1,
    u = s.lastRow - s.firstRow + 1;
  if (
    ((n.sourceRange = `${Gh(s.firstRow, s.firstCol)}:${Gh(s.lastRow, s.lastCol)}`),
    d > 50 || d * u > 500)
  )
    throw Error("Select at most 500 source cells and 50 fields.");
  if (u <= Number(n.header))
    throw Error("Select at least one data row in addition to the heading row.");
  if (
    (assertSelectionInAuthorizedSheet(a, n.sourceRange, o),
    (m = a.merges) != null &&
      m.some((v) => {
        const g = parseRangeAddress(v);
        return (
          g.firstRow <= s.lastRow &&
          g.lastRow >= s.firstRow &&
          g.firstCol <= s.lastCol &&
          g.lastCol >= s.firstCol
        );
      }))
  )
    throw Error("Unmerge the selected source before reshaping fields.");
  (assertOutputNameUnreferenced(e, n.outputName),
    discoverArrayFormulaAnchors(e, [{ sheet: a, range: n.sourceRange }], o));
  const l = [],
    c = [];
  let h = 0;
  for (let v = s.firstRow; v <= s.lastRow; v++) {
    const g = [];
    for (let b = s.firstCol; b <= s.lastCol; b++) {
      const y = Gh(v, b),
        w = a.cells[y];
      if (w && typeof w.raw == "string" && w.raw.startsWith("=") && !w.literal)
        throw Error(
          `Source ${y} is a formula. Create a reviewed values snapshot first.`,
        );
      if ((h += coerceCellValueToText(w?.raw ?? null).length) > 2e5)
        throw Error(
          "Source text exceeds the 200,000-character processing budget.",
        );
      const E = findRowOrigin(e, a.id, v);
      (l.push({
        address: y,
        ...(w ? { original: structuredClone(w) } : {}),
        ...(E ? { origin: E } : {}),
      }),
        g.push(createSourcedValue(w?.raw ?? null, [y])));
    }
    c.push({ row: v + 1, fields: g });
  }
  const f = n.header
    ? c
        .shift()
        .fields.map(
          (v, g) => coerceCellValueToText(v.value) || `Field ${g + 1}`,
        )
    : Array.from({ length: d }, (v, g) => `Field ${g + 1}`);
  if (f.some((v) => v.length > 100))
    throw Error("Source headings must be at most 100 characters.");
  return { request: n, context: o, source: a, sources: l, rows: c, headers: f };
}
function applyTransformationSteps(e, r, t, n = "1900") {
  const o = [],
    a = [],
    s = (d, u, l, c, h) => {
      const f =
        h instanceof Error
          ? h.message
          : "The transformation could not parse this value.";
      return (
        o.push({
          step: c,
          row: u,
          field: l,
          raw: d.value,
          sources: [...d.sources],
          reason: f,
        }),
        { ...d, exception: true }
      );
    };
  return (
    t.forEach((d, u) => {
      const l = u + 1,
        c = JSON.stringify(e),
        h = o.length,
        f =
          "columns" in d
            ? d.columns
            : "column" in d
              ? [d.column]
              : d.pattern.parts.map((g) => g.column);
      if ((assertRecipeFieldInRange(f, r.length), d.kind === "split")) {
        let g = d.mode === "fixed" ? d.breaks.length + 1 : 1;
        const b = e.map((w) => {
          const E = w.fields[d.column];
          if (E.exception) return [E];
          try {
            const A = splitTextWithQuoteRules(
              coerceCellValueToText(E.value),
              d,
            );
            return (
              (g = Math.max(g, A.length)),
              A.map((k) => createSourcedValue(k, [...E.sources]))
            );
          } catch (A) {
            return [s(E, w.row, r[d.column], l, A)];
          }
        });
        if (r.length - 1 + g > 50)
          throw Error("The split would exceed 50 fields.");
        e.forEach((w, E) => {
          const A = w.fields[d.column],
            k = b[E];
          for (; k.length < g;)
            k.push({
              ...createSourcedValue("", [...A.sources]),
              exception: k[0].exception,
            });
          w.fields.splice(d.column, 1, ...k);
        });
        const y = r[d.column];
        r.splice(
          d.column,
          1,
          ...Array.from({ length: g }, (w, E) => `${y.slice(0, 130)}.${E + 1}`),
        );
      } else if (d.kind === "concat") {
        if (r.length >= 50)
          throw Error("Concatenation would exceed 50 fields.");
        for (const g of e) {
          const b = d.columns.map((w) => g.fields[w]),
            y = createSourcedValue(
              b
                .filter((w) => !d.skipBlank || !isBlankCellValue(w.value))
                .map((w) => coerceCellValueToText(w.value))
                .join(d.separator),
              uniqueSourceReferences(b),
            );
          ((y.exception = b.some((w) => w.exception)), g.fields.push(y));
        }
        r.push(d.name);
      } else if (d.kind === "select") {
        for (const g of e) g.fields = d.columns.map((b) => g.fields[b]);
        r.splice(0, r.length, ...d.columns.map((g) => r[g]));
      } else if (d.kind === "fill") {
        assertRecipeFieldInRange(d.groupColumns, r.length);
        const g = d.direction === "up" ? [...e].reverse() : e;
        for (const b of d.columns) {
          let y,
            w = "";
          for (const E of g) {
            const A = d.groupColumns.map((I) => E.fields[I]),
              k = JSON.stringify(A.map((I) => I.value));
            (k !== w || A.some((I) => I.exception)) &&
              ((y = undefined), (w = k));
            const F = E.fields[b];
            if (F.exception) {
              y = undefined;
              continue;
            }
            if (!isBlankCellValue(F.value)) {
              y = F;
              continue;
            }
            d.direction === "value"
              ? (E.fields[b] = createSourcedValue(d.value, [...F.sources]))
              : y &&
                (E.fields[b] = {
                  ...y,
                  sources: uniqueSourceReferences([F, y, ...A]),
                });
          }
        }
      } else if (d.kind === "pattern") {
        if (
          (assertRecipeFieldInRange(d.inferenceColumns, r.length),
          d.pattern.parts.some((b) => !d.inferenceColumns.includes(b.column)))
        )
          throw Error(
            "The inferred pattern reads a field outside its approved inference columns.",
          );
        if (r.length >= 50)
          throw Error("The inferred output would exceed 50 fields.");
        const g = new Set();
        for (const b of d.examples) {
          const y = e.find((w) => w.row === b.row);
          if (!y || g.has(b.row))
            throw Error(
              "An inference example is outside the data rows or duplicated.",
            );
          if (
            (g.add(b.row),
            d.pattern.parts.some((w) => y.fields[w.column].exception) ||
              applyInferredTextPattern(
                y.fields.map((w) => w.value),
                d.pattern,
              ) !== b.value)
          )
            throw Error(
              "The inferred pattern no longer matches every reviewed example.",
            );
        }
        for (const b of e) {
          const y = d.pattern.parts.map((E) => b.fields[E.column]),
            w = { ...y[0], sources: uniqueSourceReferences(y) };
          try {
            if (y.some((E) => E.exception))
              throw Error("An inferred input has an earlier parse exception.");
            b.fields.push(
              createSourcedValue(
                applyInferredTextPattern(
                  b.fields.map((E) => E.value),
                  d.pattern,
                ),
                w.sources,
              ),
            );
          } catch (E) {
            b.fields.push(s(w, b.row, d.name, l, E));
          }
        }
        r.push(d.name);
      } else
        for (const g of e)
          for (const b of d.columns) {
            const y = g.fields[b];
            if (!y.exception)
              try {
                let w = y.value,
                  E = y.type;
                if (d.kind === "convert")
                  ((w = convertFieldToTypedValue(y.value, d, n)),
                    (E =
                      w === null
                        ? "blank"
                        : d.to === "date"
                          ? "date"
                          : inferCellValueType(w)));
                else if (!isBlankCellValue(y.value)) {
                  const A = coerceCellValueToText(y.value);
                  ((w =
                    d.kind === "trim"
                      ? d.mode === "collapse"
                        ? A.trim().replace(/\s+/gu, " ")
                        : A.trim()
                      : d.kind === "clean"
                        ? A.replace(/[\u0000-\u001f]/g, "")
                        : d.kind === "case"
                          ? applyTextCasing(A, d.mode)
                          : A.split(d.find).join(d.replacement)),
                    (E = "text"));
                }
                if (typeof w == "string" && w.length > 32767)
                  throw Error(
                    "The transformed value exceeds the cell text limit.",
                  );
                g.fields[b] = { ...y, value: w, type: E };
              } catch (w) {
                g.fields[b] = s(y, g.row, r[b], l, w);
              }
          }
      if (e.length * r.length > 5e3)
        throw Error("The intermediate result exceeds 5,000 cells.");
      let p = 0;
      for (const g of e)
        for (const b of g.fields) {
          if (typeof b.value == "string" && b.value.length > 32767)
            throw Error("A result exceeds the 32,767-character cell limit.");
          if ((p += coerceCellValueToText(b.value).length) > 1e6)
            throw Error(
              "Intermediate results exceed the 1,000,000-character limit.",
            );
        }
      const m = JSON.parse(c);
      let v = 0;
      (e.forEach((g, b) =>
        g.fields.forEach((y, w) => {
          JSON.stringify(y) !== JSON.stringify(m[b].fields[w]) && v++;
        }),
      ),
        a.push({
          step: l,
          kind: d.kind,
          fields: [...r],
          changedCells: v,
          exceptions: o.length - h,
        }));
    }),
    { exceptions: o, stages: a }
  );
}
function createAnnotatedCellValue(e, r, t) {
  return normalizeCellRecordInput({
    raw: e,
    note: r,
    ...(t === "date" ? { format: "date" } : {}),
    ...(typeof e == "string"
      ? { numFmt: "@", ...(e.startsWith("=") ? { literal: true } : {}) }
      : {}),
  });
}
function buildTransformationPreview(e, r, t = {}) {
  const { request, context, source, sources, rows, headers } =
    prepareTransformationInput(e, r, t);
  if (!request.steps.length)
    throw Error("Add at least one transformation step before previewing.");
  const { exceptions, stages } = applyTransformationSteps(
      rows,
      headers,
      request.steps,
      e.datePolicy?.dateSystem,
    ),
    h = [],
    f = {};
  if (
    (headers.forEach((v, g) => {
      f[Gh(0, g)] = { raw: v, literal: true, style: { bold: true } };
    }),
    rows.forEach((v, g) =>
      v.fields.forEach((b, y) => {
        const w = Gh(g + 1, y);
        h.push({ address: w, row: v.row, field: headers[y], ...b });
        const E = `Clean & Reshape derivation
Workbook: ${JSON.stringify(e.id)}
Source revision: ${e.revision}
Worksheet: ${JSON.stringify(source.name)} (${JSON.stringify(source.id)})
Original cells: ${b.sources.join(", ")}
Result type: ${b.type}
${b.exception ? "Parse exception retained; see the exceptions appendix." : "Original source values are retained."}`;
        f[w] = createAnnotatedCellValue(b.value, E, b.type);
      }),
    ),
    exceptions.length)
  ) {
    const v = rows.length + 3;
    ((f[Gh(v, 0)] = {
      raw: "Transformation exceptions",
      style: { bold: true },
    }),
      [
        "Source row",
        "Field",
        "Step",
        "Raw value",
        "Original cells",
        "Reason",
      ].forEach(
        (b, y) => (f[Gh(v + 1, y)] = { raw: b, style: { bold: true } }),
      ),
      exceptions.forEach((b, y) =>
        [
          b.row,
          b.field,
          b.step,
          b.raw,
          b.sources.map((w) => `${source.name}!${w}`).join(", "),
          b.reason,
        ].forEach(
          (w, E) =>
            (f[Gh(v + 2 + y, E)] = createAnnotatedCellValue(
              w,
              `Retained parsing exception at source revision ${e.revision}`,
              inferCellValueType(w),
            )),
        ),
      ));
  }
  if (Object.keys(f).length > 5e3)
    throw Error("Results and exception appendix exceed 5,000 output cells.");
  if (
    e.sheets.reduce((v, g) => v + Object.keys(g.cells).length, 0) +
      Object.keys(f).length >
    workbookLimits.maxCells
  )
    throw Error(
      "Derived output exceeds the 100,000 stored-cell workbook budget.",
    );
  const p = {
    snapshot: JSON.stringify(e),
    context,
    request,
    sourceSheetName: source.name,
    sources,
    headers,
    outputs: h,
    exceptions,
    stages,
    cells: f,
    targetRange: `A1:${Gh(exceptions.length ? rows.length + 4 + exceptions.length : rows.length, Math.max(headers.length, exceptions.length ? 6 : 0) - 1)}`,
    counts: {
      sourceRows: rows.length,
      sourceCells: sources.length,
      outputCells: Object.keys(f).length,
      changedCells: h.filter((v) => {
        return (
          v.sources.length !== 1 ||
          JSON.stringify(
            sources.find((y) => y.address === v.sources[0])?.original?.raw ??
              null,
          ) !== JSON.stringify(v.value)
        );
      }).length,
      exceptions: exceptions.length,
      exceptionRows: new Set(exceptions.map((v) => v.row)).size,
    },
  };
  return (
    validateFieldTransformEvidence(
      [
        ...(e.fieldTransforms ?? []),
        buildTransformationRecord(
          e,
          p,
          "preview-id",
          "preview-output",
          "2026-01-01T00:00:00.000Z",
        ),
      ],
      {},
    ),
    p
  );
}
function buildTransformationRecord(e, r, t, n, o) {
  return {
    id: t,
    workbookId: e.id,
    sourceRevision: e.revision,
    sourceSheetId: r.request.sourceSheetId,
    sourceSheetName: r.sourceSheetName,
    outputSheetId: n,
    outputSheetName: r.request.outputName,
    createdAt: o,
    semantics: "reviewed-fields-v1",
    request: structuredClone(r.request),
    sources: structuredClone(r.sources),
    outputs: structuredClone(r.outputs),
    exceptions: structuredClone(r.exceptions),
  };
}
function commitApprovedTransformation(e, r, t) {
  if (t.approved !== true)
    throw Error(
      "Review and approve the complete transformation preview first.",
    );
  if (r.snapshot !== JSON.stringify(e))
    throw Error("The workbook changed. Preview the transformation again.");
  const n = buildTransformationPreview(e, r.request, t);
  if (JSON.stringify(n) !== JSON.stringify(r))
    throw Error("The preview or authorized context changed. Preview again.");
  if (r.exceptions.length && t.acceptExceptions !== true)
    throw Error("Acknowledge the retained parse exceptions before applying.");
  const o = structuredClone(e),
    a = createWorksheetRecord(r.request.outputName);
  return (
    (a.cells = structuredClone(r.cells)),
    o.sheets.push(a),
    (o.fieldTransforms = [
      ...(o.fieldTransforms ?? []),
      buildTransformationRecord(
        e,
        r,
        crypto.randomUUID(),
        a.id,
        new Date().toISOString(),
      ),
    ]),
    normalizeWorkbookInput(o)
  );
}
function selectExceptionFreeRows(e, r, t, n, o = {}) {
  const a = prepareTransformationInput(e, r, o);
  applyTransformationSteps(
    a.rows,
    a.headers,
    a.request.steps,
    e.datePolicy?.dateSystem,
  );
  const s = a.rows.filter((u) =>
    n.every((l) => u.fields[l] && !u.fields[l].exception),
  );
  return inferTextPatternFromExamples(
    s.map((u) => ({ row: u.row, values: u.fields.map((l) => l.value) })),
    n,
    t,
  );
}
function verifyTransformationPreview(e, r, t = {}) {
  if (
    r.snapshot !== JSON.stringify(e) ||
    JSON.stringify(buildTransformationPreview(e, r.request, t)) !==
      JSON.stringify(r)
  )
    throw Error(
      "The transformation preview is stale or altered. Preview again.",
    );
  const { snapshot, context, cells, ...s } = r;
  return JSON.stringify(
    {
      schemaVersion: 1,
      kind: "field-transformation-preview",
      workbookId: e.id,
      sourceRevision: e.revision,
      ...s,
    },
    null,
    2,
  );
}
const expressionFunctionArity = {
    ABS: [1, 1],
    ROUND: [1, 2],
    UPPER: [1, 1],
    LOWER: [1, 1],
    TRIM: [1, 1],
    LEN: [1, 1],
    CONCAT: [1, 20],
    COALESCE: [1, 20],
    IF: [3, 3],
    AND: [1, 20],
    OR: [1, 20],
    NOT: [1, 1],
  },
  assertFiniteNumericOperand = (e) => {
    if (typeof e != "number" || !Number.isFinite(e))
      throw Error(
        "This expression requires a finite number. Add an explicit Types step.",
      );
    return e;
  },
  assertTextOperand = (e) => {
    if (typeof e != "string")
      throw Error("This expression requires text. Add an explicit Types step.");
    return e;
  },
  assertBooleanCondition = (e) => {
    if (typeof e != "boolean")
      throw Error("IF/AND/OR/NOT conditions must be Boolean comparisons.");
    return e;
  },
  assertDerivedValueLimits = (e) => {
    if (
      (typeof e == "number" && !Number.isFinite(e)) ||
      (typeof e == "string" && e.length > 32767)
    )
      throw Error("Derived value exceeds finite numeric or text limits.");
    return e;
  };
function parseRowExpression(e, r) {
  if (typeof e != "string" || !e.trim() || e.length > 1e3)
    throw Error("Use a row expression of 1–1,000 characters.");
  let t = 0,
    n = 0,
    o = 0;
  const a = e.trim().replace(/^=/, ""),
    s = () => {
      for (; /\s/.test(a[t] ?? "") && t < a.length;) t++;
    },
    d = () => (s(), a[t]),
    u = () => {
      throw Error(
        `Invalid row expression near character ${t + 1}. Use [Field], literals, operators and supported functions.`,
      );
    },
    l = (v) => {
      if (++n > 200) throw Error("Row expressions are limited to 200 nodes.");
      return v;
    };
  function c() {
    if (++o > 32) throw Error("Row expression nesting exceeds 32 levels.");
    try {
      let v = d();
      if (v === "+" || v === "-")
        return (t++, l({ kind: "unary", operator: v, value: c() }));
      if (v === "(") {
        t++;
        const k = f(0);
        return d() !== ")" ? u() : (t++, k);
      }
      if (v === '"') {
        const k = a.slice(t).match(/^"(?:[^"\\]|\\.)*"/);
        if (!k) return u();
        t += k[0].length;
        try {
          return l({ kind: "value", value: JSON.parse(k[0]) });
        } catch {
          return u();
        }
      }
      if (v === "[") {
        t++;
        let k = "",
          F = false;
        for (; t < a.length;) {
          if (a[t] === "]") {
            if (a[t + 1] === "]") {
              ((k += "]"), (t += 2));
              continue;
            }
            (t++, (F = true));
            break;
          }
          k += a[t++];
        }
        const I = r.indexOf(k);
        if (!F || I < 0) throw Error(`Unknown row-expression field [${k}].`);
        return l({ kind: "field", index: I });
      }
      const g = a.slice(t).match(/^(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?/);
      if (g)
        return (
          (t += g[0].length),
          l({ kind: "value", value: assertDerivedValueLimits(Number(g[0])) })
        );
      const b = a.slice(t).match(/^[A-Za-z_][A-Za-z0-9_]*/);
      if (!b) return u();
      t += b[0].length;
      const y = b[0].toUpperCase();
      if (d() !== "(") {
        if (y === "TRUE" || y === "FALSE" || y === "NULL")
          return l({
            kind: "value",
            value: y === "NULL" ? null : y === "TRUE",
          });
        throw Error(`Unknown identifier ${y}. Column references use [Field].`);
      }
      if (!Object.hasOwn(expressionFunctionArity, y))
        throw Error(
          `Unsupported row function ${y}. Use ${Object.keys(expressionFunctionArity).join(", ")}.`,
        );
      t++;
      const w = [];
      if (d() !== ")") for (; w.push(f(0)), d() === ",";) t++;
      if (d() !== ")") return u();
      t++;
      const [E, A] = expressionFunctionArity[y];
      if (w.length < E || w.length > A)
        throw Error(`${y} expects ${E === A ? E : `${E}–${A}`} arguments.`);
      return l({ kind: "call", name: y, args: w });
    } finally {
      o--;
    }
  }
  const h = {
    "=": 1,
    "<>": 1,
    "<": 1,
    ">": 1,
    "<=": 1,
    ">=": 1,
    "&": 2,
    "+": 3,
    "-": 3,
    "*": 4,
    "/": 4,
  };
  function f(v) {
    let g = c();
    for (;;) {
      s();
      const y = a.slice(t).match(/^(<>|<=|>=|[=<>+*/&-])/)?.[0];
      if (!y || h[y] < v) break;
      ((t += y.length),
        (g = l({ kind: "binary", operator: y, left: g, right: f(h[y] + 1) })));
    }
    return g;
  }
  const p = f(0);
  if ((s(), t !== a.length)) return u();
  function m(v, g) {
    if (v.kind === "value") return v.value;
    if (v.kind === "field") return g[v.index];
    if (v.kind === "unary")
      return assertDerivedValueLimits(
        assertFiniteNumericOperand(m(v.value, g)) *
          (v.operator === "-" ? -1 : 1),
      );
    if (v.kind === "binary") {
      const w = m(v.left, g),
        E = m(v.right, g),
        A = v.operator;
      if (A === "=") return w === E;
      if (A === "<>") return w !== E;
      if (["<", ">", "<=", ">="].includes(A)) {
        if (
          typeof w != typeof E ||
          w === null ||
          E === null ||
          typeof w == "boolean"
        )
          throw Error(
            "Comparisons require two numbers or two text values of the same type.",
          );
        return A === "<"
          ? w < E
          : A === ">"
            ? w > E
            : A === "<="
              ? w <= E
              : w >= E;
      }
      if (A === "&")
        return assertDerivedValueLimits(
          assertTextOperand(w) + assertTextOperand(E),
        );
      const k = assertFiniteNumericOperand(w),
        F = assertFiniteNumericOperand(E);
      if (A === "/" && F === 0)
        throw Error("Derived expression divides by zero.");
      return assertDerivedValueLimits(
        A === "+" ? k + F : A === "-" ? k - F : A === "*" ? k * F : k / F,
      );
    }
    if (v.name === "IF")
      return m(v.args[assertBooleanCondition(m(v.args[0], g)) ? 1 : 2], g);
    if (v.name === "COALESCE") {
      for (const w of v.args) {
        const E = m(w, g);
        if (E !== null && E !== "") return E;
      }
      return null;
    }
    const b = v.args.map((w) => m(w, g)),
      y = b[0];
    switch (v.name) {
      case "ABS":
        return Math.abs(assertFiniteNumericOperand(y));
      case "ROUND": {
        const w = b.length === 1 ? 0 : assertFiniteNumericOperand(b[1]);
        if (!Number.isInteger(w) || Math.abs(w) > 12)
          throw Error("ROUND digits must be an integer from -12 to 12.");
        const E = assertFiniteNumericOperand(y),
          A = 10 ** w;
        return assertDerivedValueLimits(
          (Math.sign(E) * Math.round(Math.abs(E) * A)) / A,
        );
      }
      case "UPPER":
        return assertDerivedValueLimits(
          assertTextOperand(y).toLocaleUpperCase("en-US"),
        );
      case "LOWER":
        return assertDerivedValueLimits(
          assertTextOperand(y).toLocaleLowerCase("en-US"),
        );
      case "TRIM":
        return assertTextOperand(y).trim().replace(/\s+/gu, " ");
      case "LEN":
        return Array.from(assertTextOperand(y)).length;
      case "CONCAT":
        return assertDerivedValueLimits(b.map(assertTextOperand).join(""));
      case "AND":
        return b.map(assertBooleanCondition).every(Boolean);
      case "OR":
        return b.map(assertBooleanCondition).some(Boolean);
      case "NOT":
        return !assertBooleanCondition(y);
      default:
        throw Error("Unknown row expression operation.");
    }
  }
  return (v) => assertDerivedValueLimits(m(p, v));
}
const isNullOrEmptyValue = (e) => e === null || e === "",
  fingerprintTypedValues = (e) => JSON.stringify(e.map((r) => [typeof r, r])),
  findFieldIndexOrThrow = (e, r) => {
    const t = e.fields.indexOf(r);
    if (t < 0)
      throw Error(`Field “${r}” is missing from the step input schema.`);
    return t;
  },
  uniqueSortedSources = (e) => [...new Set(e.flatMap((r) => r.sources))].sort(),
  profileFieldTypes = (e) =>
    e.fields.map((r, t) => ({
      name: r,
      types: [
        ...new Set(
          e.rows.map((n) =>
            isNullOrEmptyValue(n.values[t]) ? "blank" : typeof n.values[t],
          ),
        ),
      ].sort(),
      ...(Object.hasOwn(e.declaredTypes, r)
        ? { declaredType: e.declaredTypes[r] }
        : {}),
    }));
function buildFieldControlTotals(e) {
  return e.fields.map((r, t) => {
    let n = 0,
      o = 0,
      a = 0,
      s = 0;
    for (const d of e.rows) {
      const u = d.values[t];
      if ((isNullOrEmptyValue(u) && s++, typeof u == "number")) {
        a++;
        const l = u - o,
          c = n + l;
        if (
          ((o = c - n - l), (n = c), !Number.isFinite(n) || !Number.isFinite(o))
        )
          throw Error(
            `Control total for “${r}” exceeds finite numeric limits.`,
          );
      }
    }
    return { field: r, numericCount: a, blankCount: s, sum: n };
  });
}
function validateRecipeOutputTable(e) {
  if (
    !e.fields.length ||
    e.fields.length > transformationRecipeLimits.columns ||
    e.fields.some((n) => !n.trim() || n.length > 128) ||
    new Set(e.fields.map((n) => n.toLowerCase())).size !== e.fields.length
  )
    throw Error(
      "The result needs 1–50 distinct, nonblank column names of at most 128 characters.",
    );
  if (
    e.rows.length > transformationRecipeLimits.rows ||
    (e.rows.length + 1) * e.fields.length > transformationRecipeLimits.cells
  )
    throw Error(
      "A recipe step exceeds 1,000 rows or 20,000 cells. Narrow the source; no rows were truncated.",
    );
  let r = 0,
    t = 0;
  for (const n of e.rows) {
    if (n.values.length !== e.fields.length)
      throw Error("Invalid recipe row width.");
    t += n.sources.length;
    for (const o of n.values)
      if (typeof o == "string") {
        if (((r += o.length), o.length > 32767))
          throw Error("Recipe text exceeds the cell limit.");
      } else if (
        o !== null &&
        typeof o != "boolean" &&
        (typeof o != "number" || !Number.isFinite(o))
      )
        throw Error("Recipe output must contain finite scalar values.");
  }
  if (
    r > transformationRecipeLimits.characters ||
    t > transformationRecipeLimits.lineageLinks
  )
    throw Error(
      "Recipe text or source-lineage budget exceeded. Narrow the source; no evidence was truncated.",
    );
  return e;
}
function resolveRecipeSourceTable(e, r, t) {
  const n = e.sheets.find((c) => c.id === r.sheetId);
  if (!n || (t.allowedSheetIds && !t.allowedSheetIds.includes(n.id)))
    throw Error("A recipe source is outside the authorized host context.");
  let o = r.range;
  if (r.tableId) {
    const c = n.tables?.find((f) => f.id === r.tableId);
    if (!c)
      throw Error(
        "The selected source table no longer exists. Review the recipe source.",
      );
    const h = parseRangeAddress(c.range);
    o = `${columnIndexToLetters(h.firstCol)}${h.firstRow + 1}:${columnIndexToLetters(h.lastCol)}${h.lastRow + 1 - (c.showTotals ? 1 : 0)}`;
  }
  const a = readSourceHeaders(e, [{ sheetId: r.sheetId, range: o }], t),
    s = a.snapshots[0],
    d = validateRecipeOutputTable({
      fields: s.headers,
      rows: s.records.map((c) => ({ values: c.values, sources: [c.locator] })),
      declaredTypes: {},
    }),
    u = {
      source: r,
      resolvedRange: s.source.range,
      sheetName: n.name,
      hash: fingerprintCanonicalValue({
        workbookId: e.id,
        source: r,
        snapshot: s,
        dependencies: a.dependencyCells,
        datePolicy: normalizeDatePolicy(e.datePolicy),
      }),
      schema: profileFieldTypes(d),
      rows: d.rows.length,
      controls: buildFieldControlTotals(d),
    };
  return { data: d, evidence: u };
}
function aggregateFieldValues(e, r, t) {
  const n = e.map((s) => s.values[r]).filter((s) => !isNullOrEmptyValue(s));
  if (t === "count") return n.length;
  if (t === "distinctCount")
    return new Set(n.map((s) => fingerprintTypedValues([s]))).size;
  if (n.some((s) => typeof s != "number"))
    throw Error(
      "Numeric aggregation encountered text or Boolean values. Add a Types step before aggregating.",
    );
  const o = n;
  if (!o.length) return t === "sum" ? 0 : null;
  let a =
    t === "min"
      ? Math.min(...o)
      : t === "max"
        ? Math.max(...o)
        : o.reduce((s, d) => s + d, 0);
  if ((t === "average" && (a /= o.length), !Number.isFinite(a)))
    throw Error("Aggregation exceeds finite numeric bounds.");
  return a;
}
function groupRowsByFields(e, r) {
  const t = r.map((o) => findFieldIndexOrThrow(e, o)),
    n = new Map();
  for (const o of e.rows) {
    const a = fingerprintTypedValues(t.map((s) => o.values[s]));
    n.set(a, [...(n.get(a) ?? []), o]);
  }
  return { indices: t, groups: n };
}
function executeRecipeStep(e, r, t, n) {
  const o = r,
    a = (s) => s.map((d) => findFieldIndexOrThrow(o, d));
  switch (t.kind) {
    case "select": {
      const s = a(t.fields);
      return {
        data: {
          fields: t.fields,
          rows: o.rows.map((d) => ({
            ...d,
            values: s.map((u) => d.values[u]),
          })),
          declaredTypes: Object.fromEntries(
            t.fields.flatMap((d) =>
              Object.hasOwn(o.declaredTypes, d)
                ? [[d, o.declaredTypes[d]]]
                : [],
            ),
          ),
        },
      };
    }
    case "rename": {
      if (
        (a(t.mappings.map((d) => d.from)),
        new Set(t.mappings.map((d) => d.from)).size !== t.mappings.length)
      )
        throw Error("Rename each input field only once per step.");
      const s = (d) => {
        return t.mappings.find((l) => l.from === d)?.to ?? d;
      };
      return {
        data: {
          ...o,
          fields: o.fields.map(s),
          declaredTypes: Object.fromEntries(
            Object.entries(o.declaredTypes).map(([d, u]) => [s(d), u]),
          ),
        },
      };
    }
    case "types": {
      const s = a(t.fields);
      if (t.group && t.group === t.decimal)
        throw Error("Decimal and grouping separators must differ.");
      return {
        data: {
          ...o,
          declaredTypes: {
            ...o.declaredTypes,
            ...Object.fromEntries(t.fields.map((d) => [d, t.to])),
          },
          rows: o.rows.map((d, u) => {
            try {
              return {
                ...d,
                values: d.values.map((l, c) =>
                  s.includes(c)
                    ? convertFieldToTypedValue(
                        l,
                        {
                          kind: "convert",
                          columns: s,
                          to: t.to,
                          decimal: t.decimal,
                          group: t.group,
                          dateOrder: t.dateOrder,
                        },
                        normalizeDatePolicy(e.datePolicy).dateSystem,
                      )
                    : l,
                ),
              };
            } catch (l) {
              throw Error(
                `Input record ${u + 1} (${d.sources.join(", ")}): ${l.message}`,
              );
            }
          }),
        },
      };
    }
    case "replace": {
      const s = a(t.fields);
      if (
        t.mode === "text" &&
        (typeof t.find != "string" ||
          !t.find ||
          typeof t.replacement != "string")
      )
        throw Error(
          "Text replace needs nonempty text to find and a text replacement.",
        );
      const d = (u) => {
        const l = u.split(t.find);
        if (
          u.length + (l.length - 1) * (t.replacement.length - t.find.length) >
          32767
        )
          throw Error("Replacement exceeds the cell text limit.");
        return l.join(t.replacement);
      };
      return {
        data: {
          ...o,
          declaredTypes: {},
          rows: o.rows.map((u) => ({
            ...u,
            values: u.values.map((l, c) =>
              s.includes(c)
                ? t.mode === "value"
                  ? l === t.find
                    ? t.replacement
                    : l
                  : typeof l == "string"
                    ? d(l)
                    : l
                : l,
            ),
          })),
        },
      };
    }
    case "derive": {
      if (o.fields.some((d) => d.toLowerCase() === t.name.toLowerCase()))
        throw Error("The derived column name already exists.");
      const s = parseRowExpression(t.expression, o.fields);
      return {
        data: {
          ...o,
          fields: [...o.fields, t.name],
          rows: o.rows.map((d, u) => {
            try {
              return { ...d, values: [...d.values, s(d.values)] };
            } catch (l) {
              throw Error(
                `Input record ${u + 1} (${d.sources.join(", ")}): ${l.message}`,
              );
            }
          }),
        },
      };
    }
    case "filter": {
      const s = t.conditions.map((u) => {
          const l = findFieldIndexOrThrow(o, u.field);
          if (
            !["blank", "notBlank"].includes(u.operator) &&
            u.value === undefined
          )
            throw Error("This filter requires an explicit typed value.");
          if (
            ["contains", "startsWith", "endsWith"].includes(u.operator) &&
            typeof u.value != "string"
          )
            throw Error("Text filters require a text criterion.");
          if (
            ["gt", "gte", "lt", "lte"].includes(u.operator) &&
            typeof u.value != "number"
          )
            throw Error("Ordered filters require a numeric criterion.");
          return { ...u, i: l };
        }),
        d = o.rows.filter((u, l) => {
          const c = s.map((h) => {
            const f = u.values[h.i],
              p = h.value;
            switch (h.operator) {
              case "blank":
                return isNullOrEmptyValue(f);
              case "notBlank":
                return !isNullOrEmptyValue(f);
              case "equals":
                return f === p;
              case "notEquals":
                return f !== p;
              case "contains":
                return typeof f == "string" && f.includes(p);
              case "startsWith":
                return typeof f == "string" && f.startsWith(p);
              case "endsWith":
                return typeof f == "string" && f.endsWith(p);
              default:
                if (isNullOrEmptyValue(f)) return false;
                if (typeof f != "number")
                  throw Error(
                    `Ordered filter found a nonnumeric value in “${h.field}” on input record ${l + 1}. Convert types explicitly.`,
                  );
                return h.operator === "gt"
                  ? f > p
                  : h.operator === "gte"
                    ? f >= p
                    : h.operator === "lt"
                      ? f < p
                      : f <= p;
            }
          });
          return t.join === "and" ? c.every(Boolean) : c.some(Boolean);
        });
      return {
        data: { ...o, rows: d },
        details: {
          retained: d.length,
          excluded: o.rows.length - d.length,
          comparison:
            "Exact typed comparison; text uses case-sensitive Unicode code units.",
        },
      };
    }
    case "sort": {
      const s = t.fields.map((l) => ({
          ...l,
          i: findFieldIndexOrThrow(o, l.field),
        })),
        d = (l) =>
          l === null
            ? 3
            : typeof l == "number"
              ? 0
              : typeof l == "string"
                ? 1
                : 2,
        u = o.rows
          .map((l, c) => ({ r: l, i: c }))
          .sort((l, c) => {
            for (const h of s) {
              const f = l.r.values[h.i],
                p = c.r.values[h.i];
              let m = d(f) - d(p);
              if ((!m && f !== p && (m = f < p ? -1 : 1), m))
                return m * (h.direction === "ascending" ? 1 : -1);
            }
            return l.i - c.i;
          })
          .map((l) => l.r);
      return { data: { ...o, rows: u } };
    }
    case "unpivot": {
      const s = a(t.fields),
        d = o.fields.map((u, l) => l).filter((u) => !s.includes(u));
      if (o.rows.length * s.length > transformationRecipeLimits.rows)
        throw Error(
          "Unpivot would exceed 1,000 intermediate rows; narrow the selected fields or source.",
        );
      return {
        data: {
          fields: [...d.map((u) => o.fields[u]), t.attribute, t.value],
          rows: o.rows.flatMap((u) =>
            s
              .filter((l) => t.keepNulls || !isNullOrEmptyValue(u.values[l]))
              .map((l) => ({
                values: [
                  ...d.map((c) => u.values[c]),
                  o.fields[l],
                  u.values[l],
                ],
                sources: u.sources,
              })),
          ),
          declaredTypes: { [t.attribute]: "text" },
        },
        details: {
          nullPolicy: t.keepNulls ? "retain" : "exclude",
          selectedFields: t.fields,
        },
      };
    }
    case "group": {
      const { indices, groups } = groupRowsByFields(o, t.keys),
        u = t.values.map((l) => ({
          ...l,
          i: findFieldIndexOrThrow(o, l.field),
        }));
      return {
        data: {
          fields: [...t.keys, ...u.map((l) => l.name)],
          rows: [...groups.values()].map((l) => ({
            values: [
              ...indices.map((c) => l[0].values[c]),
              ...u.map((c) => aggregateFieldValues(l, c.i, c.operation)),
            ],
            sources: uniqueSortedSources(l),
          })),
          declaredTypes: {},
        },
      };
    }
    case "pivot": {
      if (
        t.keys.includes(t.column) ||
        t.keys.includes(t.value) ||
        t.column === t.value
      )
        throw Error("Pivot keys, attribute and value fields must be distinct.");
      const s = findFieldIndexOrThrow(o, t.column),
        d = findFieldIndexOrThrow(o, t.value),
        { indices, groups } = groupRowsByFields(o, t.keys),
        c = [
          ...new Map(
            o.rows.map((p) => [
              fingerprintTypedValues([p.values[s]]),
              p.values[s],
            ]),
          ).values(),
        ],
        h = c.map(
          (p) =>
            `${p === null ? "blank" : typeof p}:${p === null ? "(blank)" : String(p)}`,
        );
      if (t.keys.length + h.length > 50)
        throw Error("Pivot exceeds 50 output columns.");
      const f = [...groups.values()].map((p) => ({
        values: [
          ...indices.map((m) => p[0].values[m]),
          ...c.map((m) =>
            aggregateFieldValues(
              p.filter((v) => v.values[s] === m),
              d,
              t.operation,
            ),
          ),
        ],
        sources: uniqueSortedSources(p),
      }));
      if (!c.length && t.keys.length === 0)
        throw Error(
          "An empty pivot has no columns. Select at least one grouping key.",
        );
      return {
        data: { fields: [...t.keys, ...h], rows: f, declaredTypes: {} },
        details: {
          columns: c.map((p, m) => ({ header: h[m], value: p })),
          aggregation: t.operation,
        },
      };
    }
    case "append": {
      const s = n(t.source);
      if (
        o.fields.length !== s.fields.length ||
        o.fields.some((u) => !s.fields.includes(u))
      )
        throw Error(
          "Append requires the same exact header set; column order may differ.",
        );
      const d = o.fields.map((u) => findFieldIndexOrThrow(s, u));
      return {
        data: {
          fields: o.fields,
          rows: [
            ...o.rows,
            ...s.rows.map((u) => ({ ...u, values: d.map((l) => u.values[l]) })),
          ],
          declaredTypes: {},
        },
        details: {
          leftRows: o.rows.length,
          rightRows: s.rows.length,
          outputRows: o.rows.length + s.rows.length,
        },
      };
    }
    case "merge": {
      const s = n(t.source),
        d = t.keys.map((v) => findFieldIndexOrThrow(o, v.left)),
        u = t.keys.map((v) => findFieldIndexOrThrow(s, v.right));
      if (new Set(d).size !== d.length || new Set(u).size !== u.length)
        throw Error("Merge keys cannot repeat a field.");
      const l = (v, g) =>
          v.rows.map((b, y) => ({
            row: y,
            locator: `${g} input record ${y + 1}`,
            values: b.values,
          })),
        c = l(o, "Left"),
        h = l(s, "Right"),
        f = matchKeyPopulations(c, h, d, u, t.normalization, t.fuzzy),
        p = [],
        m = (v, g) => {
          if (p.length >= 1e3)
            throw Error("Merge exceeds 1,000 intermediate records.");
          p.push({
            values: [
              ...(v?.values ?? o.fields.map(() => null)),
              ...(g?.values ?? s.fields.map(() => null)),
            ],
            sources: uniqueSortedSources([v, g].filter((b) => !!b)),
          });
        };
      for (const v of c) {
        const g = f.byLeft.get(v.row);
        if (g) {
          if (!["leftAnti", "rightAnti"].includes(t.mode))
            for (const b of g) m(o.rows[v.row], s.rows[b.record.row]);
        } else
          ["left", "full", "leftAnti"].includes(t.mode) &&
            m(o.rows[v.row], undefined);
      }
      if (["right", "full", "rightAnti"].includes(t.mode))
        for (const v of f.unmatchedRight) m(undefined, s.rows[v.row]);
      return {
        data: {
          fields: [...o.fields, ...s.fields.map((v) => t.prefix + v)],
          rows: p,
          declaredTypes: {},
        },
        details: {
          leftRows: o.rows.length,
          rightRows: s.rows.length,
          matchedLeft: f.matchedLeftRows,
          matchedRight: f.matchedRightRows,
          matchPairs: f.matchPairs,
          unmatchedLeft: f.unmatchedLeft,
          unmatchedRight: f.unmatchedRight,
          duplicateLeft: f.duplicateLeftKeys,
          duplicateRight: f.duplicateRightKeys,
          manyToMany: f.manyToManyKeys,
          extraPairs: f.expandedMatchRows,
          matching: f.diagnostics,
        },
      };
    }
  }
}

export {
  executeRecipeStep,
  commitApprovedTransformation,
  discoverArrayFormulaAnchors,
  resolveRecipeSourceTable,
  verifyTransformationPreview,
  applyTextSplitPreview,
  selectExceptionFreeRows,
  buildTransformationPreview,
  validateRecipeOutputTable,
  buildTextSplitResult,
  describeInferredTextRule,
  profileFieldTypes,
  buildFieldControlTotals,
};
