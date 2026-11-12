import {
  isNullOrEmptyValue,
  serializeTypedKeyValues,
  failWithMessage,
} from "./03-workbook-model.js";
import { formatCellAddress } from "./04-cell-input-paste-and-fill.js";
import {
  transformationRecipeLimits,
  fingerprintCanonicalValue,
  normalizeDatePolicy,
  dateSystemMaxSerial,
  parseDateText,
} from "./01-schemas-date-functions-and-validation.js";
import { workbookLimits } from "./02-text-and-financial-functions.js";
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
} from "./12-consolidation-and-dependency-inspection.js";
import {
  validateTextToColumnsRequest,
  assertSelectionInAuthorizedSheet,
  splitDelimitedCell,
  parseImportedCellValue,
  rangesOverlap,
  dataModelBudgets,
  formatCellTextValue,
  validateTextToColumnsPermissions,
  assertOutputNameUnreferenced,
} from "./15-data-model-dialog.js";

function discoverArrayFormulaAnchors(workbook, r, context) {
  const n = [];
  let o = 0,
    a = 0;
  for (const sheet of new Map(r.map((f) => [f.sheet.id, f.sheet])).values())
    for (const [anchor, cell] of Object.entries(sheet.cells)) {
      if (
        typeof cell.raw != "string" ||
        !cell.raw.startsWith("=") ||
        cell.literal
      )
        continue;
      (++o > dataModelBudgets.formulaCells ||
        (a += cell.raw.length) > dataModelBudgets.formulaCharacters) &&
        failWithMessage(
          "Array checking supports at most 1,000 formulas and 100,000 formula characters on the participating worksheets.",
        );
      const m = _t(anchor);
      r.some((g) => {
        if (g.sheet.id !== sheet.id) return false;
        const rangeAddress = parseRangeAddress(g.range),
          y = Math.max(m.row, rangeAddress.firstRow),
          w = Math.max(m.col, rangeAddress.firstCol);
        return (
          y <= rangeAddress.lastRow &&
          w <= rangeAddress.lastCol &&
          (y - m.row + 1) * (w - m.col + 1) <= 1e4
        );
      }) && n.push({ sheet: sheet, anchor: anchor });
    }
  if (
    (n.length > dataModelBudgets.discoveryAnchors &&
      failWithMessage(
        "The selection has more than 50 possible array anchors. Choose a new output sheet or a smaller reviewed region.",
      ),
    !n.length)
  )
    return;
  const s = new Set(),
    d = (h, f) => JSON.stringify([h, f]),
    u = {
      ...workbook,
      sheets: workbook.sheets.map((sheet) =>
        sheet.veryHidden ? { ...sheet, hidden: true } : sheet,
      ),
    };
  for (const h of n) {
    assertSelectionInAuthorizedSheet(h.sheet, h.anchor, context);
    const f = inspectFormulaDependencies(u, {
        sheetId: h.sheet.id,
        address: h.anchor,
        direction: "precedents",
        depth: 8,
        includeHidden: context.includeHidden,
        allowedSheetIds: context.allowedSheetIds,
        limits: { maxFormulaCells: 1e3, maxNodes: 500, maxReferences: 2e3 },
      }),
      p = [f.target, ...f.nodes],
      m = new Set(p.map((event) => event.key));
    (f.coverage.truncated ||
      f.coverage.depthLimited ||
      f.cycles.length ||
      f.issues.some((issue) => m.has(issue.from))) &&
      failWithMessage(
        "Array discovery requires a complete, authorized static dependency closure. Resolve hidden, dynamic, cyclic or unsupported formula dependencies before splitting.",
      );
    for (const v of p) {
      const g = workbook.sheets.find((sheet) => sheet.id === v.sheetId);
      (assertSelectionInAuthorizedSheet(g, v.address, context),
        (!g.cells[v.address] || g.cells[v.address].raw === null) &&
          failWithMessage(
            "Array discovery cannot infer blank or possible spill-child dependencies. Use a reviewed values snapshot before splitting.",
          ),
        s.add(d(v.sheetId, v.address)),
        s.size > dataModelBudgets.dependencyCells &&
          failWithMessage(
            "Array discovery exceeds 2,000 dependency cells. Use a smaller reviewed values snapshot.",
          ));
    }
  }
  const l = {
      ...workbook,
      sheets: workbook.sheets
        .filter((sheet) => context.allowedSheetIds.includes(sheet.id))
        .map((h) => ({
          ...h,
          cells: Object.fromEntries(
            Object.entries(h.cells).filter(([f]) => s.has(d(h.id, f))),
          ),
        })),
    },
    workbookWithSettings = calculateWorkbookWithSettings(l);
  for (const { sheet, anchor } of n) {
    const p = workbookWithSettings.getSpillRange(sheet.id, anchor);
    p &&
      r.some((m) => m.sheet.id === sheet.id && rangesOverlap(m.range, p)) &&
      failWithMessage(
        "The selected source or destination intersects a generated array. Convert its results to ordinary values explicitly before splitting.",
      );
  }
  workbookWithSettings.getOperationalErrors().length &&
    failWithMessage(
      "Array discovery encountered an operational calculation error. Resolve the uncertain formula state before splitting text.",
    );
}
function buildTextSplitResult(workbook, request, t = {}) {
  const context = validateTextToColumnsPermissions(workbook, t),
    o = validateTextToColumnsRequest(request),
    a = workbook.sheets.find((sheet) => sheet.id === o.sourceSheetId);
  (a || failWithMessage("Choose an existing source worksheet."),
    assertSelectionInAuthorizedSheet(a, o.sourceRange, context));
  const rangeAddress = parseRangeAddress(o.sourceRange),
    height = rangeAddress.lastRow - rangeAddress.firstRow + 1;
  ((rangeAddress.firstCol !== rangeAddress.lastCol ||
    height > dataModelBudgets.sourceCells) &&
    failWithMessage("Select one source column containing at most 500 cells."),
    a.merges?.some((merge) => rangesOverlap(merge, o.sourceRange)) &&
      failWithMessage("Unmerge the source before splitting text."));
  const sources = [],
    l = [],
    diagnostics = [];
  let h = 0,
    f = 0;
  for (
    let firstRow = rangeAddress.firstRow;
    firstRow <= rangeAddress.lastRow;
    firstRow++
  ) {
    const cellAddress = formatCellAddress(firstRow, rangeAddress.firstCol),
      cell = a.cells[cellAddress];
    cell &&
      typeof cell.raw == "string" &&
      cell.raw.startsWith("=") &&
      !cell.literal &&
      failWithMessage(
        `Source ${cellAddress} contains a formula. Copy its results to ordinary values explicitly before splitting.`,
      );
    const cellTextValue = formatCellTextValue(cell?.raw ?? null);
    (f += cellTextValue.length) > dataModelBudgets.sourceCharacters &&
      failWithMessage(
        "Selected source text exceeds the 200,000-character processing budget.",
      );
    let D;
    try {
      D = splitDelimitedCell(cellTextValue, o.delimiter);
    } catch (error) {
      failWithMessage(`Source ${cellAddress}: ${error.message}`);
    }
    ((h = Math.max(h, D.length)),
      l.push(D),
      sources.push({
        address: cellAddress,
        ...(cell ? { cell: structuredClone(cell) } : {}),
        text: cellTextValue,
        fields: D.length,
      }));
  }
  (h * height > dataModelBudgets.outputCells &&
    failWithMessage(
      "Split output may contain at most 500 cells. Select fewer source rows or fields.",
    ),
    (o.formats?.length ?? 0) > h &&
      failWithMessage(
        "A column conversion refers to a field outside the split result.",
      ));
  const formats = Array.from({ length: h }, (O, j) => o.formats?.[j] ?? "text");
  let m,
    targetSheetName,
    g = { row: 0, col: 0 };
  if (o.destination.kind === "new-sheet")
    ((!context.allowNewSheet || workbook.structureProtection?.enabled) &&
      failWithMessage(
        "Creating an output worksheet is not permitted while workbook structure is protected or the host disallows new worksheets.",
      ),
      (targetSheetName = o.destination.name),
      workbook.sheets.some(
        (sheet) => sheet.name.toLowerCase() === targetSheetName.toLowerCase(),
      ) &&
        failWithMessage(
          "The output worksheet name is already used. Choose another name.",
        ),
      assertOutputNameUnreferenced(workbook, targetSheetName));
  else {
    const destination = o.destination;
    ((m = workbook.sheets.find((sheet) => sheet.id === destination.sheetId)),
      m || failWithMessage("Choose an existing output worksheet."),
      (g = _t(o.destination.address)),
      (targetSheetName = m.name),
      m.protection?.enabled &&
        failWithMessage("The output worksheet is protected."));
  }
  (g.row + height > workbookLimits.maxRows ||
    g.col + h > workbookLimits.maxCols) &&
    failWithMessage("Split output exceeds the worksheet boundary.");
  const b = `${formatCellAddress(g.row, g.col)}:${formatCellAddress(g.row + height - 1, g.col + h - 1)}`;
  (m &&
    (assertSelectionInAuthorizedSheet(m, b, context),
    m.id === a.id &&
      rangesOverlap(o.sourceRange, b) &&
      failWithMessage(
        "Choose a separate destination so the original source cells remain intact.",
      ),
    (m.tables?.some((table) => rangesOverlap(table.range, b)) ||
      m.merges?.some((merge) => rangesOverlap(merge, b)) ||
      [...(m.pivots ?? []), ...(m.queries ?? [])].some((O) =>
        rangesOverlap(O.outputRange, b),
      )) &&
      failWithMessage(
        "Choose an output outside tables, merged areas and generated analysis results.",
      ),
    Object.keys(m.cells).some((cell) => rangesOverlap(cell, b)) &&
      failWithMessage(
        "The destination contains cells or their formatting. Choose an empty range or new worksheet.",
      ),
    [
      ...(workbook.reviewThreads ?? []),
      ...(workbook.formulaWatches ?? []),
    ].some(
      (O) =>
        !O.orphaned &&
        O.anchor.sheetId === m.id &&
        rangesOverlap(O.anchor.address, b),
    ) &&
      failWithMessage(
        "The destination contains a saved watch or discussion anchor. Choose another range to preserve its evidence context.",
      )),
    discoverArrayFormulaAnchors(
      workbook,
      [
        { sheet: a, range: o.sourceRange },
        ...(m ? [{ sheet: m, range: b }] : []),
      ],
      context,
    ));
  const outputs = [];
  for (let O = 0; O < height; O++) {
    l[O].length !== h &&
      diagnostics.push({
        code: "padded-row",
        sourceAddress: sources[O].address,
        message: `This record has ${l[O].length} fields; ${h - l[O].length} missing fields are padded with empty text.`,
      });
    for (let j = 0; j < h; j++) {
      const field = l[O][j] ?? "",
        format = o.header === "first-row" && O === 0 ? "text" : formats[j],
        importedCellValue = parseImportedCellValue(field, format);
      importedCellValue.precision &&
        diagnostics.push({
          code: "numeric-text-retained",
          sourceAddress: sources[O].address,
          message: `Field ${j + 1} remains text because numeric conversion would exceed the supported finite precision.`,
        });
      const note = `Text to Columns derivation
Workbook ID: ${JSON.stringify(workbook.id)}
Source revision: ${workbook.revision}
Worksheet ID: ${JSON.stringify(a.id)}
Worksheet name: ${JSON.stringify(a.name)}
Source cell: ${sources[O].address}
Field: ${j + 1}
Conversion: ${format}
${j >= l[O].length ? "Padded missing field." : "Original source cell and evidence remain in place."}`,
        cellRecordInput = normalizeCellRecordInput({
          raw: importedCellValue.raw,
          note: note,
          ...(typeof importedCellValue.raw == "string"
            ? {
                numFmt: "@",
                ...(importedCellValue.raw.startsWith("=")
                  ? { literal: true }
                  : {}),
              }
            : {}),
          ...(o.header === "first-row" && O === 0
            ? { style: { bold: true } }
            : {}),
        });
      outputs.push({
        sourceAddress: sources[O].address,
        address: formatCellAddress(g.row + O, g.col + j),
        field: field,
        padded: j >= l[O].length,
        format: format,
        type:
          typeof importedCellValue.raw == "number"
            ? "number"
            : typeof importedCellValue.raw == "boolean"
              ? "boolean"
              : "text",
        cell: cellRecordInput,
      });
    }
  }
  if (
    (workbook.sheets.reduce((O, j) => O + Object.keys(j.cells).length, 0) +
      outputs.length >
      workbookLimits.maxCells &&
      failWithMessage(
        "The split would exceed the workbook's 100,000 stored-cell budget.",
      ),
    m)
  ) {
    const book = structuredClone(workbook),
      j = book.sheets.find((sheet) => sheet.id === m.id);
    for (const output of outputs) j.cells[output.address] = output.cell;
    checkDataValidationBatch(
      book,
      outputs.map((output) => ({
        sheetId: j.id,
        address: output.address,
        after: output.cell,
      })),
    );
  }
  return {
    snapshot: JSON.stringify(workbook),
    request: o,
    context: context,
    sourceSheetName: a.name,
    targetSheetName: targetSheetName,
    targetRange: b,
    width: h,
    height: height,
    formats: formats,
    sources: sources,
    outputs: outputs,
    diagnostics: diagnostics,
  };
}
function applyTextSplitPreview(workbook, r, t = {}) {
  r.snapshot !== JSON.stringify(workbook) &&
    failWithMessage(
      "The workbook changed after the split preview. Preview again.",
    );
  const textSplitResult = buildTextSplitResult(workbook, r.request, t);
  JSON.stringify(textSplitResult) !== JSON.stringify(r) &&
    failWithMessage(
      "The split preview or authorized context changed. Preview again.",
    );
  const o = structuredClone(workbook),
    destination = r.request.destination,
    s =
      destination.kind === "new-sheet"
        ? createWorksheetRecord(destination.name)
        : o.sheets.find((sheet) => sheet.id === destination.sheetId);
  destination.kind === "new-sheet" && o.sheets.push(s);
  for (const output of r.outputs)
    s.cells[output.address] = structuredClone(output.cell);
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
function extractInferredFieldPart(e, part) {
  if (part.column >= e.length)
    throw Error("The inferred rule refers to a missing field.");
  const t = coerceCellValueToText(e[part.column]).trim();
  if (part.mode === "whole") return t;
  if (part.mode === "slice") {
    const s = [...t];
    if (s.length < part.start + part.length)
      throw Error(
        "The value is too short for the inferred character extraction.",
      );
    return s.slice(part.start, part.start + part.length).join("");
  }
  const n = t.indexOf(part.delimiter);
  if (part.mode === "before" || part.mode === "after") {
    if (n < 0) throw Error("The inferred separator is absent.");
    return (
      part.mode === "before"
        ? t.slice(0, n)
        : t.slice(n + part.delimiter.length)
    ).trim();
  }
  const o = part.delimiter === " " ? t.split(/ +/) : t.split(part.delimiter),
    a = part.index < 0 ? o.length + part.index : part.index;
  if (a < 0 || a >= o.length || !o[a])
    throw Error("The inferred token is absent.");
  return o[a].trim();
}
function applyInferredTextPattern(e, pattern) {
  const t = textJoinSpecSchema.parse(pattern),
    n =
      t.prefix +
      applyTextCasing(
        t.parts
          .map((part) => extractInferredFieldPart(e, part))
          .join(t.separator),
        t.casing,
      ) +
      t.suffix;
  if (n.length > 32767)
    throw Error("The inferred result exceeds the cell text limit.");
  return n;
}
function describeInferredTextRule(pattern) {
  const r = pattern.parts
    .map(
      (part) =>
        `field ${part.column + 1}${part.mode === "whole" ? "" : part.mode === "token" ? ` token ${part.index < 0 ? "last" : part.index + 1} by ${JSON.stringify(part.delimiter)}` : part.mode === "slice" ? ` characters ${part.start + 1}–${part.start + part.length}` : ` ${part.mode} ${JSON.stringify(part.delimiter)}`}`,
    )
    .join(` + ${JSON.stringify(pattern.separator)} + `);
  return `${pattern.prefix ? JSON.stringify(pattern.prefix) + " + " : ""}${pattern.casing === "none" ? "" : pattern.casing + "("}${r}${pattern.casing === "none" ? "" : ")"}${pattern.suffix ? " + " + JSON.stringify(pattern.suffix) : ""}`;
}
function inferTextPatternFromExamples(e, r, t) {
  if (!Array.isArray(t) || t.length < 2 || t.length > 10)
    throw Error("Provide 2–10 examples from distinct source rows.");
  const n = t.map((f) => sourceRowValueSchema.parse(f));
  if (new Set(n.map((f) => f.row)).size !== n.length)
    throw Error("Provide examples from distinct source rows.");
  const o = n.map((example) => ({
    example: example,
    row: e.find((p) => p.row === example.row),
  }));
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
    s = (column) => ({
      column: column,
      mode: "whole",
      delimiter: " ",
      index: 0,
      start: 0,
      length: 1,
    });
  for (const f of r) {
    a.push(s(f));
    for (const delimiter of [" ", ",", "-", "/", "_", "@"]) {
      for (const m of [0, 1, -1])
        a.push({ ...s(f), mode: "token", delimiter: delimiter, index: m });
      for (const mode of ["before", "after"])
        a.push({ ...s(f), mode: mode, delimiter: delimiter });
    }
    for (const p of [1, 2, 3, 4, 5, 6, 8, 10])
      a.push({ ...s(f), mode: "slice", length: p });
  }
  const d = a.map((f) => ({ parts: [f], separator: "" }));
  for (const f of r)
    for (const p of r)
      if (f !== p)
        for (const separator of [" ", "", ", ", ".", "-", "_", "/"])
          d.push({ parts: [s(f), s(p)], separator: separator });
  const candidates = [],
    l = new Set();
  let searched = 0,
    truncated = false;
  e: for (const f of d)
    for (const casing of ["none", "upper", "lower", "proper"]) {
      if (++searched > textInferenceLimits.inferencePatterns) {
        (searched--, (truncated = true));
        break e;
      }
      const m = { ...f, casing: casing, prefix: "", suffix: "" };
      let v;
      try {
        v = o.map((E) => applyInferredTextPattern(E.row.values, m));
      } catch {
        continue;
      }
      if (v.some((E) => !E) || new Set(v).size < 2) continue;
      const value = o[0].example.value,
        b = v[0];
      let y = value.indexOf(b),
        w = 0;
      for (; y >= 0 && w++ < 3;) {
        const pattern = {
          ...m,
          prefix: value.slice(0, y),
          suffix: value.slice(y + b.length),
        };
        if (
          pattern.prefix.length <= 1e3 &&
          pattern.suffix.length <= 1e3 &&
          o.every(
            (A, k) =>
              pattern.prefix + v[k] + pattern.suffix === A.example.value,
          )
        ) {
          const A = JSON.stringify(pattern);
          if (
            !l.has(A) &&
            (l.add(A),
            candidates.push({
              id: `pattern-${candidates.length + 1}`,
              pattern: pattern,
              description: describeInferredTextRule(pattern),
              matchedExamples: n.length,
            }),
            candidates.length >= textInferenceLimits.inferenceCandidates)
          ) {
            truncated = true;
            break e;
          }
        }
        y = value.indexOf(b, y + 1);
      }
    }
  return { candidates: candidates, searched: searched, truncated: truncated };
}
const inferCellValueType = (e) =>
    e === null
      ? "blank"
      : typeof e == "number"
        ? "number"
        : typeof e == "boolean"
          ? "boolean"
          : "text",
  uniqueSourceReferences = (e) => [...new Set(e.flatMap((r) => r.sources))],
  createSourcedValue = (e, sources, t = inferCellValueType(e)) => ({
    value: e,
    type: t,
    sources: sources,
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
function convertFieldToTypedValue(value, r, dateSystem = "1900") {
  if (r.to === "text") return coerceCellValueToText(value);
  if (isNullOrEmptyValue(value)) return null;
  if (r.to === "boolean") {
    if (typeof value == "boolean") return value;
    if (typeof value == "string" && /^(true|false)$/i.test(value.trim()))
      return value.trim().toUpperCase() === "TRUE";
    throw Error("Expected TRUE or FALSE.");
  }
  if (r.to === "date") {
    if (
      typeof value == "number" &&
      Number.isInteger(value) &&
      value >= (dateSystem === "1904" ? 0 : 1) &&
      value <= dateSystemMaxSerial(dateSystem)
    )
      return value;
    if (
      typeof value != "string" ||
      !/^\s*(?:\d{4}-\d{1,2}-\d{1,2}|\d{1,2}\/\d{1,2}\/\d{4})\s*$/.test(value)
    )
      throw Error(
        "Expected a full-year ISO date or an explicitly ordered slash date.",
      );
    const dateText = parseDateText(value, {
      textDateOrder: r.dateOrder,
      dateSystem: dateSystem,
    });
    if (dateText === undefined)
      throw Error(`The date is impossible or outside ${dateSystem}–9999.`);
    return dateText;
  }
  if (typeof value == "boolean")
    throw Error("Boolean values are not numeric fields.");
  let n;
  if (typeof value == "number") n = value;
  else {
    let d = value.trim();
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
function prepareTransformationInput(workbook, r, t) {
  const request = validateFieldTransformRecipe(r),
    context = validateTextToColumnsPermissions(workbook, t),
    a = workbook.sheets.find((sheet) => sheet.id === request.sourceSheetId);
  if (!a) throw Error("Choose an existing source worksheet.");
  if (!context.allowNewSheet || workbook.structureProtection?.enabled)
    throw Error("Creating a derived worksheet is not permitted.");
  if (
    workbook.sheets.some(
      (sheet) => sheet.name.toLowerCase() === request.outputName.toLowerCase(),
    )
  )
    throw Error("The output worksheet name is already used.");
  const rangeAddress = parseRangeAddress(request.sourceRange),
    d = rangeAddress.lastCol - rangeAddress.firstCol + 1,
    u = rangeAddress.lastRow - rangeAddress.firstRow + 1;
  if (
    ((request.sourceRange = `${formatCellAddress(rangeAddress.firstRow, rangeAddress.firstCol)}:${formatCellAddress(rangeAddress.lastRow, rangeAddress.lastCol)}`),
    d > 50 || d * u > 500)
  )
    throw Error("Select at most 500 source cells and 50 fields.");
  if (u <= Number(request.header))
    throw Error("Select at least one data row in addition to the heading row.");
  if (
    (assertSelectionInAuthorizedSheet(a, request.sourceRange, context),
    a.merges?.some((merge) => {
      const g = parseRangeAddress(merge);
      return (
        g.firstRow <= rangeAddress.lastRow &&
        g.lastRow >= rangeAddress.firstRow &&
        g.firstCol <= rangeAddress.lastCol &&
        g.lastCol >= rangeAddress.firstCol
      );
    }))
  )
    throw Error("Unmerge the selected source before reshaping fields.");
  (assertOutputNameUnreferenced(workbook, request.outputName),
    discoverArrayFormulaAnchors(
      workbook,
      [{ sheet: a, range: request.sourceRange }],
      context,
    ));
  const sources = [],
    rows = [];
  let h = 0;
  for (
    let firstRow = rangeAddress.firstRow;
    firstRow <= rangeAddress.lastRow;
    firstRow++
  ) {
    const fields = [];
    for (
      let firstCol = rangeAddress.firstCol;
      firstCol <= rangeAddress.lastCol;
      firstCol++
    ) {
      const cellAddress = formatCellAddress(firstRow, firstCol),
        cell = a.cells[cellAddress];
      if (
        cell &&
        typeof cell.raw == "string" &&
        cell.raw.startsWith("=") &&
        !cell.literal
      )
        throw Error(
          `Source ${cellAddress} is a formula. Create a reviewed values snapshot first.`,
        );
      if ((h += coerceCellValueToText(cell?.raw ?? null).length) > 2e5)
        throw Error(
          "Source text exceeds the 200,000-character processing budget.",
        );
      const rowOrigin = findRowOrigin(workbook, a.id, firstRow);
      (sources.push({
        address: cellAddress,
        ...(cell ? { original: structuredClone(cell) } : {}),
        ...(rowOrigin ? { origin: rowOrigin } : {}),
      }),
        fields.push(createSourcedValue(cell?.raw ?? null, [cellAddress])));
    }
    rows.push({ row: firstRow + 1, fields: fields });
  }
  const headers = request.header
    ? rows
        .shift()
        .fields.map(
          (field, g) => coerceCellValueToText(field.value) || `Field ${g + 1}`,
        )
    : Array.from({ length: d }, (v, g) => `Field ${g + 1}`);
  if (headers.some((header) => header.length > 100))
    throw Error("Source headings must be at most 100 characters.");
  return {
    request: request,
    context: context,
    source: a,
    sources: sources,
    rows: rows,
    headers: headers,
  };
}
function applyTransformationSteps(rows, headers, steps, dateSystem = "1900") {
  const exceptions = [],
    stages = [],
    s = (d, row, field, step, h) => {
      const reason =
        h instanceof Error
          ? h.message
          : "The transformation could not parse this value.";
      return (
        exceptions.push({
          step: step,
          row: row,
          field: field,
          raw: d.value,
          sources: [...d.sources],
          reason: reason,
        }),
        { ...d, exception: true }
      );
    };
  return (
    steps.forEach((d, u) => {
      const step = u + 1,
        c = JSON.stringify(rows),
        h = exceptions.length,
        f =
          "columns" in d
            ? d.columns
            : "column" in d
              ? [d.column]
              : d.pattern.parts.map((part) => part.column);
      if ((assertRecipeFieldInRange(f, headers.length), d.kind === "split")) {
        let g = d.mode === "fixed" ? d.breaks.length + 1 : 1;
        const b = rows.map((row) => {
          const E = row.fields[d.column];
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
          } catch (error) {
            return [s(E, row.row, headers[d.column], step, error)];
          }
        });
        if (headers.length - 1 + g > 50)
          throw Error("The split would exceed 50 fields.");
        rows.forEach((row, E) => {
          const A = row.fields[d.column],
            k = b[E];
          for (; k.length < g;)
            k.push({
              ...createSourcedValue("", [...A.sources]),
              exception: k[0].exception,
            });
          row.fields.splice(d.column, 1, ...k);
        });
        const y = headers[d.column];
        headers.splice(
          d.column,
          1,
          ...Array.from({ length: g }, (w, E) => `${y.slice(0, 130)}.${E + 1}`),
        );
      } else if (d.kind === "concat") {
        if (headers.length >= 50)
          throw Error("Concatenation would exceed 50 fields.");
        for (const row of rows) {
          const b = d.columns.map((column) => row.fields[column]),
            sourcedValue = createSourcedValue(
              b
                .filter((w) => !d.skipBlank || !isNullOrEmptyValue(w.value))
                .map((w) => coerceCellValueToText(w.value))
                .join(d.separator),
              uniqueSourceReferences(b),
            );
          ((sourcedValue.exception = b.some((w) => w.exception)),
            row.fields.push(sourcedValue));
        }
        headers.push(d.name);
      } else if (d.kind === "select") {
        for (const row of rows)
          row.fields = d.columns.map((column) => row.fields[column]);
        headers.splice(
          0,
          headers.length,
          ...d.columns.map((column) => headers[column]),
        );
      } else if (d.kind === "fill") {
        assertRecipeFieldInRange(d.groupColumns, headers.length);
        const g = d.direction === "up" ? [...rows].reverse() : rows;
        for (const column of d.columns) {
          let y,
            w = "";
          for (const E of g) {
            const A = d.groupColumns.map(
                (groupColumn) => E.fields[groupColumn],
              ),
              k = JSON.stringify(A.map((I) => I.value));
            (k !== w || A.some((I) => I.exception)) &&
              ((y = undefined), (w = k));
            const F = E.fields[column];
            if (F.exception) {
              y = undefined;
              continue;
            }
            if (!isNullOrEmptyValue(F.value)) {
              y = F;
              continue;
            }
            d.direction === "value"
              ? (E.fields[column] = createSourcedValue(d.value, [...F.sources]))
              : y &&
                (E.fields[column] = {
                  ...y,
                  sources: uniqueSourceReferences([F, y, ...A]),
                });
          }
        }
      } else if (d.kind === "pattern") {
        if (
          (assertRecipeFieldInRange(d.inferenceColumns, headers.length),
          d.pattern.parts.some(
            (part) => !d.inferenceColumns.includes(part.column),
          ))
        )
          throw Error(
            "The inferred pattern reads a field outside its approved inference columns.",
          );
        if (headers.length >= 50)
          throw Error("The inferred output would exceed 50 fields.");
        const g = new Set();
        for (const example of d.examples) {
          const row = rows.find((w) => w.row === example.row);
          if (!row || g.has(example.row))
            throw Error(
              "An inference example is outside the data rows or duplicated.",
            );
          if (
            (g.add(example.row),
            d.pattern.parts.some((part) => row.fields[part.column].exception) ||
              applyInferredTextPattern(
                row.fields.map((field) => field.value),
                d.pattern,
              ) !== example.value)
          )
            throw Error(
              "The inferred pattern no longer matches every reviewed example.",
            );
        }
        for (const row of rows) {
          const y = d.pattern.parts.map((part) => row.fields[part.column]),
            w = { ...y[0], sources: uniqueSourceReferences(y) };
          try {
            if (y.some((E) => E.exception))
              throw Error("An inferred input has an earlier parse exception.");
            row.fields.push(
              createSourcedValue(
                applyInferredTextPattern(
                  row.fields.map((field) => field.value),
                  d.pattern,
                ),
                w.sources,
              ),
            );
          } catch (error) {
            row.fields.push(s(w, row.row, d.name, step, error));
          }
        }
        headers.push(d.name);
      } else
        for (const row of rows)
          for (const column of d.columns) {
            const y = row.fields[column];
            if (!y.exception)
              try {
                let value = y.value,
                  type = y.type;
                if (d.kind === "convert")
                  ((value = convertFieldToTypedValue(y.value, d, dateSystem)),
                    (type =
                      value === null
                        ? "blank"
                        : d.to === "date"
                          ? "date"
                          : inferCellValueType(value)));
                else if (!isNullOrEmptyValue(y.value)) {
                  const A = coerceCellValueToText(y.value);
                  ((value =
                    d.kind === "trim"
                      ? d.mode === "collapse"
                        ? A.trim().replace(/\s+/gu, " ")
                        : A.trim()
                      : d.kind === "clean"
                        ? A.replace(/[\u0000-\u001f]/g, "")
                        : d.kind === "case"
                          ? applyTextCasing(A, d.mode)
                          : A.split(d.find).join(d.replacement)),
                    (type = "text"));
                }
                if (typeof value == "string" && value.length > 32767)
                  throw Error(
                    "The transformed value exceeds the cell text limit.",
                  );
                row.fields[column] = { ...y, value: value, type: type };
              } catch (error) {
                row.fields[column] = s(
                  y,
                  row.row,
                  headers[column],
                  step,
                  error,
                );
              }
          }
      if (rows.length * headers.length > 5e3)
        throw Error("The intermediate result exceeds 5,000 cells.");
      let p = 0;
      for (const row of rows)
        for (const field of row.fields) {
          if (typeof field.value == "string" && field.value.length > 32767)
            throw Error("A result exceeds the 32,767-character cell limit.");
          if ((p += coerceCellValueToText(field.value).length) > 1e6)
            throw Error(
              "Intermediate results exceed the 1,000,000-character limit.",
            );
        }
      const m = JSON.parse(c);
      let changedCells = 0;
      (rows.forEach((row, b) =>
        row.fields.forEach((field, w) => {
          JSON.stringify(field) !== JSON.stringify(m[b].fields[w]) &&
            changedCells++;
        }),
      ),
        stages.push({
          step: step,
          kind: d.kind,
          fields: [...headers],
          changedCells: changedCells,
          exceptions: exceptions.length - h,
        }));
    }),
    { exceptions: exceptions, stages: stages }
  );
}
function createAnnotatedCellValue(raw, note, t) {
  return normalizeCellRecordInput({
    raw: raw,
    note: note,
    ...(t === "date" ? { format: "date" } : {}),
    ...(typeof raw == "string"
      ? { numFmt: "@", ...(raw.startsWith("=") ? { literal: true } : {}) }
      : {}),
  });
}
function buildTransformationPreview(workbook, r, t = {}) {
  const { request, context, source, sources, rows, headers } =
    prepareTransformationInput(workbook, r, t);
  if (!request.steps.length)
    throw Error("Add at least one transformation step before previewing.");
  const { exceptions, stages } = applyTransformationSteps(
      rows,
      headers,
      request.steps,
      workbook.datePolicy?.dateSystem,
    ),
    outputs = [],
    cells = {};
  if (
    (headers.forEach((header, g) => {
      cells[formatCellAddress(0, g)] = {
        raw: header,
        literal: true,
        style: { bold: true },
      };
    }),
    rows.forEach((row, g) =>
      row.fields.forEach((field, y) => {
        const cellAddress = formatCellAddress(g + 1, y);
        outputs.push({
          address: cellAddress,
          row: row.row,
          field: headers[y],
          ...field,
        });
        const note = `Clean & Reshape derivation
Workbook: ${JSON.stringify(workbook.id)}
Source revision: ${workbook.revision}
Worksheet: ${JSON.stringify(source.name)} (${JSON.stringify(source.id)})
Original cells: ${field.sources.join(", ")}
Result type: ${field.type}
${field.exception ? "Parse exception retained; see the exceptions appendix." : "Original source values are retained."}`;
        cells[cellAddress] = createAnnotatedCellValue(
          field.value,
          note,
          field.type,
        );
      }),
    ),
    exceptions.length)
  ) {
    const v = rows.length + 3;
    ((cells[formatCellAddress(v, 0)] = {
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
        (raw, y) =>
          (cells[formatCellAddress(v + 1, y)] = {
            raw: raw,
            style: { bold: true },
          }),
      ),
      exceptions.forEach((exception, y) =>
        [
          exception.row,
          exception.field,
          exception.step,
          exception.raw,
          exception.sources.map((w) => `${source.name}!${w}`).join(", "),
          exception.reason,
        ].forEach(
          (raw, E) =>
            (cells[formatCellAddress(v + 2 + y, E)] = createAnnotatedCellValue(
              raw,
              `Retained parsing exception at source revision ${workbook.revision}`,
              inferCellValueType(raw),
            )),
        ),
      ));
  }
  if (Object.keys(cells).length > 5e3)
    throw Error("Results and exception appendix exceed 5,000 output cells.");
  if (
    workbook.sheets.reduce((v, g) => v + Object.keys(g.cells).length, 0) +
      Object.keys(cells).length >
    workbookLimits.maxCells
  )
    throw Error(
      "Derived output exceeds the 100,000 stored-cell workbook budget.",
    );
  const p = {
    snapshot: JSON.stringify(workbook),
    context,
    request,
    sourceSheetName: source.name,
    sources,
    headers,
    outputs: outputs,
    exceptions,
    stages,
    cells: cells,
    targetRange: `A1:${formatCellAddress(exceptions.length ? rows.length + 4 + exceptions.length : rows.length, Math.max(headers.length, exceptions.length ? 6 : 0) - 1)}`,
    counts: {
      sourceRows: rows.length,
      sourceCells: sources.length,
      outputCells: Object.keys(cells).length,
      changedCells: outputs.filter(
        (output) =>
          output.sources.length !== 1 ||
          JSON.stringify(
            sources.find((y) => y.address === output.sources[0])?.original
              ?.raw ?? null,
          ) !== JSON.stringify(output.value),
      ).length,
      exceptions: exceptions.length,
      exceptionRows: new Set(exceptions.map((exception) => exception.row)).size,
    },
  };
  return (
    validateFieldTransformEvidence(
      [
        ...(workbook.fieldTransforms ?? []),
        buildTransformationRecord(
          workbook,
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
function buildTransformationRecord(workbook, r, t, outputSheetId, createdAt) {
  return {
    id: t,
    workbookId: workbook.id,
    sourceRevision: workbook.revision,
    sourceSheetId: r.request.sourceSheetId,
    sourceSheetName: r.sourceSheetName,
    outputSheetId: outputSheetId,
    outputSheetName: r.request.outputName,
    createdAt: createdAt,
    semantics: "reviewed-fields-v1",
    request: structuredClone(r.request),
    sources: structuredClone(r.sources),
    outputs: structuredClone(r.outputs),
    exceptions: structuredClone(r.exceptions),
  };
}
function commitApprovedTransformation(workbook, r, t) {
  if (t.approved !== true)
    throw Error(
      "Review and approve the complete transformation preview first.",
    );
  if (r.snapshot !== JSON.stringify(workbook))
    throw Error("The workbook changed. Preview the transformation again.");
  const transformationPreview = buildTransformationPreview(
    workbook,
    r.request,
    t,
  );
  if (JSON.stringify(transformationPreview) !== JSON.stringify(r))
    throw Error("The preview or authorized context changed. Preview again.");
  if (r.exceptions.length && t.acceptExceptions !== true)
    throw Error("Acknowledge the retained parse exceptions before applying.");
  const o = structuredClone(workbook),
    worksheetRecord = createWorksheetRecord(r.request.outputName);
  return (
    (worksheetRecord.cells = structuredClone(r.cells)),
    o.sheets.push(worksheetRecord),
    (o.fieldTransforms = [
      ...(o.fieldTransforms ?? []),
      buildTransformationRecord(
        workbook,
        r,
        crypto.randomUUID(),
        worksheetRecord.id,
        new Date().toISOString(),
      ),
    ]),
    normalizeWorkbookInput(o)
  );
}
function selectExceptionFreeRows(workbook, r, t, n, o = {}) {
  const a = prepareTransformationInput(workbook, r, o);
  applyTransformationSteps(
    a.rows,
    a.headers,
    a.request.steps,
    workbook.datePolicy?.dateSystem,
  );
  const s = a.rows.filter((row) =>
    n.every((l) => row.fields[l] && !row.fields[l].exception),
  );
  return inferTextPatternFromExamples(
    s.map((u) => ({
      row: u.row,
      values: u.fields.map((field) => field.value),
    })),
    n,
    t,
  );
}
function verifyTransformationPreview(workbook, r, t = {}) {
  if (
    r.snapshot !== JSON.stringify(workbook) ||
    JSON.stringify(buildTransformationPreview(workbook, r.request, t)) !==
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
      workbookId: workbook.id,
      sourceRevision: workbook.revision,
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
function parseRowExpression(expression, fields) {
  if (
    typeof expression != "string" ||
    !expression.trim() ||
    expression.length > 1e3
  )
    throw Error("Use a row expression of 1–1,000 characters.");
  let t = 0,
    n = 0,
    o = 0;
  const a = expression.trim().replace(/^=/, ""),
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
      let operator = d();
      if (operator === "+" || operator === "-")
        return (t++, l({ kind: "unary", operator: operator, value: c() }));
      if (operator === "(") {
        t++;
        const k = f(0);
        return d() !== ")" ? u() : (t++, k);
      }
      if (operator === '"') {
        const k = a.slice(t).match(/^"(?:[^"\\]|\\.)*"/);
        if (!k) return u();
        t += k[0].length;
        try {
          return l({ kind: "value", value: JSON.parse(k[0]) });
        } catch {
          return u();
        }
      }
      if (operator === "[") {
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
        const I = fields.indexOf(k);
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
      const args = [];
      if (d() !== ")") for (; args.push(f(0)), d() === ",";) t++;
      if (d() !== ")") return u();
      t++;
      const [E, A] = expressionFunctionArity[y];
      if (args.length < E || args.length > A)
        throw Error(`${y} expects ${E === A ? E : `${E}–${A}`} arguments.`);
      return l({ kind: "call", name: y, args: args });
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
    let left = c();
    for (;;) {
      s();
      const operator = a.slice(t).match(/^(<>|<=|>=|[=<>+*/&-])/)?.[0];
      if (!operator || h[operator] < v) break;
      ((t += operator.length),
        (left = l({
          kind: "binary",
          operator: operator,
          left: left,
          right: f(h[operator] + 1),
        })));
    }
    return left;
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
        operator = v.operator;
      if (operator === "=") return w === E;
      if (operator === "<>") return w !== E;
      if (["<", ">", "<=", ">="].includes(operator)) {
        if (
          typeof w != typeof E ||
          w === null ||
          E === null ||
          typeof w == "boolean"
        )
          throw Error(
            "Comparisons require two numbers or two text values of the same type.",
          );
        return operator === "<"
          ? w < E
          : operator === ">"
            ? w > E
            : operator === "<="
              ? w <= E
              : w >= E;
      }
      if (operator === "&")
        return assertDerivedValueLimits(
          assertTextOperand(w) + assertTextOperand(E),
        );
      const k = assertFiniteNumericOperand(w),
        F = assertFiniteNumericOperand(E);
      if (operator === "/" && F === 0)
        throw Error("Derived expression divides by zero.");
      return assertDerivedValueLimits(
        operator === "+"
          ? k + F
          : operator === "-"
            ? k - F
            : operator === "*"
              ? k * F
              : k / F,
      );
    }
    if (v.name === "IF")
      return m(v.args[assertBooleanCondition(m(v.args[0], g)) ? 1 : 2], g);
    if (v.name === "COALESCE") {
      for (const arg of v.args) {
        const E = m(arg, g);
        if (E !== null && E !== "") return E;
      }
      return null;
    }
    const b = v.args.map((arg) => m(arg, g)),
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
const findFieldIndexOrThrow = (e, r) => {
    const t = e.fields.indexOf(r);
    if (t < 0)
      throw Error(`Field “${r}” is missing from the step input schema.`);
    return t;
  },
  uniqueSortedSources = (e) => [...new Set(e.flatMap((r) => r.sources))].sort(),
  profileFieldTypes = (e) =>
    e.fields.map((field, t) => ({
      name: field,
      types: [
        ...new Set(
          e.rows.map((row) =>
            isNullOrEmptyValue(row.values[t]) ? "blank" : typeof row.values[t],
          ),
        ),
      ].sort(),
      ...(Object.hasOwn(e.declaredTypes, field)
        ? { declaredType: e.declaredTypes[field] }
        : {}),
    }));
function buildFieldControlTotals(output) {
  return output.fields.map((field, t) => {
    let sum = 0,
      o = 0,
      numericCount = 0,
      blankCount = 0;
    for (const row of output.rows) {
      const u = row.values[t];
      if ((isNullOrEmptyValue(u) && blankCount++, typeof u == "number")) {
        numericCount++;
        const l = u - o,
          c = sum + l;
        if (
          ((o = c - sum - l),
          (sum = c),
          !Number.isFinite(sum) || !Number.isFinite(o))
        )
          throw Error(
            `Control total for “${field}” exceeds finite numeric limits.`,
          );
      }
    }
    return {
      field: field,
      numericCount: numericCount,
      blankCount: blankCount,
      sum: sum,
    };
  });
}
function validateRecipeOutputTable(e) {
  if (
    !e.fields.length ||
    e.fields.length > transformationRecipeLimits.columns ||
    e.fields.some((field) => !field.trim() || field.length > 128) ||
    new Set(e.fields.map((field) => field.toLowerCase())).size !==
      e.fields.length
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
  for (const row of e.rows) {
    if (row.values.length !== e.fields.length)
      throw Error("Invalid recipe row width.");
    t += row.sources.length;
    for (const value of row.values)
      if (typeof value == "string") {
        if (((r += value.length), value.length > 32767))
          throw Error("Recipe text exceeds the cell limit.");
      } else if (
        value !== null &&
        typeof value != "boolean" &&
        (typeof value != "number" || !Number.isFinite(value))
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
function resolveRecipeSourceTable(workbook, source, t) {
  const n = workbook.sheets.find((sheet) => sheet.id === source.sheetId);
  if (!n || (t.allowedSheetIds && !t.allowedSheetIds.includes(n.id)))
    throw Error("A recipe source is outside the authorized host context.");
  let range = source.range;
  if (source.tableId) {
    const c = n.tables?.find((table) => table.id === source.tableId);
    if (!c)
      throw Error(
        "The selected source table no longer exists. Review the recipe source.",
      );
    const rangeAddress = parseRangeAddress(c.range);
    range = `${columnIndexToLetters(rangeAddress.firstCol)}${rangeAddress.firstRow + 1}:${columnIndexToLetters(rangeAddress.lastCol)}${rangeAddress.lastRow + 1 - (c.showTotals ? 1 : 0)}`;
  }
  const sourceHeaders = readSourceHeaders(
      workbook,
      [{ sheetId: source.sheetId, range: range }],
      t,
    ),
    snapshot = sourceHeaders.snapshots[0],
    output = validateRecipeOutputTable({
      fields: snapshot.headers,
      rows: snapshot.records.map((record) => ({
        values: record.values,
        sources: [record.locator],
      })),
      declaredTypes: {},
    }),
    evidence = {
      source: source,
      resolvedRange: snapshot.source.range,
      sheetName: n.name,
      hash: fingerprintCanonicalValue({
        workbookId: workbook.id,
        source: source,
        snapshot: snapshot,
        dependencies: sourceHeaders.dependencyCells,
        datePolicy: normalizeDatePolicy(workbook.datePolicy),
      }),
      schema: profileFieldTypes(output),
      rows: output.rows.length,
      controls: buildFieldControlTotals(output),
    };
  return { data: output, evidence: evidence };
}
function aggregateFieldValues(e, r, operation) {
  const n = e.map((s) => s.values[r]).filter((s) => !isNullOrEmptyValue(s));
  if (operation === "count") return n.length;
  if (operation === "distinctCount")
    return new Set(n.map((s) => serializeTypedKeyValues([s]))).size;
  if (n.some((s) => typeof s != "number"))
    throw Error(
      "Numeric aggregation encountered text or Boolean values. Add a Types step before aggregating.",
    );
  const o = n;
  if (!o.length) return operation === "sum" ? 0 : null;
  let a =
    operation === "min"
      ? Math.min(...o)
      : operation === "max"
        ? Math.max(...o)
        : o.reduce((s, d) => s + d, 0);
  if ((operation === "average" && (a /= o.length), !Number.isFinite(a)))
    throw Error("Aggregation exceeds finite numeric bounds.");
  return a;
}
function groupRowsByFields(e, keys) {
  const indices = keys.map((key) => findFieldIndexOrThrow(e, key)),
    groups = new Map();
  for (const row of e.rows) {
    const a = serializeTypedKeyValues(
      indices.map((indice) => row.values[indice]),
    );
    groups.set(a, [...(groups.get(a) ?? []), row]);
  }
  return { indices: indices, groups: groups };
}
function executeRecipeStep(workbook, output, t, n) {
  const o = output,
    a = (s) => s.map((d) => findFieldIndexOrThrow(o, d));
  switch (t.kind) {
    case "select": {
      const s = a(t.fields);
      return {
        data: {
          fields: t.fields,
          rows: o.rows.map((row) => ({
            ...row,
            values: s.map((u) => row.values[u]),
          })),
          declaredTypes: Object.fromEntries(
            t.fields.flatMap((field) =>
              Object.hasOwn(o.declaredTypes, field)
                ? [[field, o.declaredTypes[field]]]
                : [],
            ),
          ),
        },
      };
    }
    case "rename": {
      if (
        (a(t.mappings.map((mapping) => mapping.from)),
        new Set(t.mappings.map((mapping) => mapping.from)).size !==
          t.mappings.length)
      )
        throw Error("Rename each input field only once per step.");
      const s = (d) =>
        t.mappings.find((mapping) => mapping.from === d)?.to ?? d;
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
      const columns = a(t.fields);
      if (t.group && t.group === t.decimal)
        throw Error("Decimal and grouping separators must differ.");
      return {
        data: {
          ...o,
          declaredTypes: {
            ...o.declaredTypes,
            ...Object.fromEntries(t.fields.map((field) => [field, t.to])),
          },
          rows: o.rows.map((row, u) => {
            try {
              return {
                ...row,
                values: row.values.map((value, c) =>
                  columns.includes(c)
                    ? convertFieldToTypedValue(
                        value,
                        {
                          kind: "convert",
                          columns: columns,
                          to: t.to,
                          decimal: t.decimal,
                          group: t.group,
                          dateOrder: t.dateOrder,
                        },
                        normalizeDatePolicy(workbook.datePolicy).dateSystem,
                      )
                    : value,
                ),
              };
            } catch (error) {
              throw Error(
                `Input record ${u + 1} (${row.sources.join(", ")}): ${error.message}`,
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
          rows: o.rows.map((row) => ({
            ...row,
            values: row.values.map((value, c) =>
              s.includes(c)
                ? t.mode === "value"
                  ? value === t.find
                    ? t.replacement
                    : value
                  : typeof value == "string"
                    ? d(value)
                    : value
                : value,
            ),
          })),
        },
      };
    }
    case "derive": {
      if (
        o.fields.some((field) => field.toLowerCase() === t.name.toLowerCase())
      )
        throw Error("The derived column name already exists.");
      const rowExpression = parseRowExpression(t.expression, o.fields);
      return {
        data: {
          ...o,
          fields: [...o.fields, t.name],
          rows: o.rows.map((row, u) => {
            try {
              return {
                ...row,
                values: [...row.values, rowExpression(row.values)],
              };
            } catch (error) {
              throw Error(
                `Input record ${u + 1} (${row.sources.join(", ")}): ${error.message}`,
              );
            }
          }),
        },
      };
    }
    case "filter": {
      const s = t.conditions.map((condition) => {
          const fieldIndexOrThrow = findFieldIndexOrThrow(o, condition.field);
          if (
            !["blank", "notBlank"].includes(condition.operator) &&
            condition.value === undefined
          )
            throw Error("This filter requires an explicit typed value.");
          if (
            ["contains", "startsWith", "endsWith"].includes(
              condition.operator,
            ) &&
            typeof condition.value != "string"
          )
            throw Error("Text filters require a text criterion.");
          if (
            ["gt", "gte", "lt", "lte"].includes(condition.operator) &&
            typeof condition.value != "number"
          )
            throw Error("Ordered filters require a numeric criterion.");
          return { ...condition, i: fieldIndexOrThrow };
        }),
        rows = o.rows.filter((row, l) => {
          const c = s.map((h) => {
            const f = row.values[h.i],
              value = h.value;
            switch (h.operator) {
              case "blank":
                return isNullOrEmptyValue(f);
              case "notBlank":
                return !isNullOrEmptyValue(f);
              case "equals":
                return f === value;
              case "notEquals":
                return f !== value;
              case "contains":
                return typeof f == "string" && f.includes(value);
              case "startsWith":
                return typeof f == "string" && f.startsWith(value);
              case "endsWith":
                return typeof f == "string" && f.endsWith(value);
              default:
                if (isNullOrEmptyValue(f)) return false;
                if (typeof f != "number")
                  throw Error(
                    `Ordered filter found a nonnumeric value in “${h.field}” on input record ${l + 1}. Convert types explicitly.`,
                  );
                return h.operator === "gt"
                  ? f > value
                  : h.operator === "gte"
                    ? f >= value
                    : h.operator === "lt"
                      ? f < value
                      : f <= value;
            }
          });
          return t.join === "and" ? c.every(Boolean) : c.some(Boolean);
        });
      return {
        data: { ...o, rows: rows },
        details: {
          retained: rows.length,
          excluded: o.rows.length - rows.length,
          comparison:
            "Exact typed comparison; text uses case-sensitive Unicode code units.",
        },
      };
    }
    case "sort": {
      const s = t.fields.map((field) => ({
          ...field,
          i: findFieldIndexOrThrow(o, field.field),
        })),
        d = (l) =>
          l === null
            ? 3
            : typeof l == "number"
              ? 0
              : typeof l == "string"
                ? 1
                : 2,
        rows = o.rows
          .map((row, c) => ({ r: row, i: c }))
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
      return { data: { ...o, rows: rows } };
    }
    case "unpivot": {
      const s = a(t.fields),
        d = o.fields.map((field, l) => l).filter((u) => !s.includes(u));
      if (o.rows.length * s.length > transformationRecipeLimits.rows)
        throw Error(
          "Unpivot would exceed 1,000 intermediate rows; narrow the selected fields or source.",
        );
      return {
        data: {
          fields: [...d.map((u) => o.fields[u]), t.attribute, t.value],
          rows: o.rows.flatMap((row) =>
            s
              .filter((l) => t.keepNulls || !isNullOrEmptyValue(row.values[l]))
              .map((l) => ({
                values: [
                  ...d.map((c) => row.values[c]),
                  o.fields[l],
                  row.values[l],
                ],
                sources: row.sources,
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
        u = t.values.map((value) => ({
          ...value,
          i: findFieldIndexOrThrow(o, value.field),
        }));
      return {
        data: {
          fields: [...t.keys, ...u.map((l) => l.name)],
          rows: [...groups.values()].map((l) => ({
            values: [
              ...indices.map((indice) => l[0].values[indice]),
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
      const fieldIndexOrThrow = findFieldIndexOrThrow(o, t.column),
        d = findFieldIndexOrThrow(o, t.value),
        { indices, groups } = groupRowsByFields(o, t.keys),
        c = [
          ...new Map(
            o.rows.map((row) => [
              serializeTypedKeyValues([row.values[fieldIndexOrThrow]]),
              row.values[fieldIndexOrThrow],
            ]),
          ).values(),
        ],
        h = c.map(
          (p) =>
            `${p === null ? "blank" : typeof p}:${p === null ? "(blank)" : String(p)}`,
        );
      if (t.keys.length + h.length > 50)
        throw Error("Pivot exceeds 50 output columns.");
      const rows = [...groups.values()].map((p) => ({
        values: [
          ...indices.map((indice) => p[0].values[indice]),
          ...c.map((m) =>
            aggregateFieldValues(
              p.filter((v) => v.values[fieldIndexOrThrow] === m),
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
        data: { fields: [...t.keys, ...h], rows: rows, declaredTypes: {} },
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
        o.fields.some((field) => !s.fields.includes(field))
      )
        throw Error(
          "Append requires the same exact header set; column order may differ.",
        );
      const d = o.fields.map((field) => findFieldIndexOrThrow(s, field));
      return {
        data: {
          fields: o.fields,
          rows: [
            ...o.rows,
            ...s.rows.map((row) => ({
              ...row,
              values: d.map((l) => row.values[l]),
            })),
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
        d = t.keys.map((key) => findFieldIndexOrThrow(o, key.left)),
        u = t.keys.map((key) => findFieldIndexOrThrow(s, key.right));
      if (new Set(d).size !== d.length || new Set(u).size !== u.length)
        throw Error("Merge keys cannot repeat a field.");
      const l = (v, g) =>
          v.rows.map((row, y) => ({
            row: y,
            locator: `${g} input record ${y + 1}`,
            values: row.values,
          })),
        records = l(o, "Left"),
        h = l(s, "Right"),
        f = matchKeyPopulations(records, h, d, u, t.normalization, t.fuzzy),
        rows = [],
        m = (v, g) => {
          if (rows.length >= 1e3)
            throw Error("Merge exceeds 1,000 intermediate records.");
          rows.push({
            values: [
              ...(v?.values ?? o.fields.map(() => null)),
              ...(g?.values ?? s.fields.map(() => null)),
            ],
            sources: uniqueSortedSources([v, g].filter((b) => !!b)),
          });
        };
      for (const record of records) {
        const g = f.byLeft.get(record.row);
        if (g) {
          if (!["leftAnti", "rightAnti"].includes(t.mode))
            for (const b of g) m(o.rows[record.row], s.rows[b.record.row]);
        } else
          ["left", "full", "leftAnti"].includes(t.mode) &&
            m(o.rows[record.row], undefined);
      }
      if (["right", "full", "rightAnti"].includes(t.mode))
        for (const v of f.unmatchedRight) m(undefined, s.rows[v.row]);
      return {
        data: {
          fields: [...o.fields, ...s.fields.map((field) => t.prefix + field)],
          rows: rows,
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
export { coerceCellValueToText };
