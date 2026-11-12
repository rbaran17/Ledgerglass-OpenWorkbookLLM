import { parseDelimitedCellValue } from "./06c-package-and-worksheet-parsing.js";

import {
  F$,
  Hle,
  Wle,
  buildContainsTextFormula,
  buildDelimitedImportReport,
  buildExchangeSchemaReport,
  byteCharacterTable,
  computeCrc32,
  convertTemplateToXlsx,
  createEmptyExportLedger,
  createWorkbookParseState,
  dataValidationOperators,
  decodeXmlEntities,
  delimitedImportDefaults,
  formatCellAddress,
  formatRangeCornerCells,
  getXmlAttribute,
  inspectXlsxPackage,
  isExternalActionFormula,
  isExternalReferenceFormula,
  maxXlsxExpandedByteLimit,
  normalizeDelimitedOptions,
  numericPrecisionPolicy,
  parseCalculationProperties,
  parseWorksheetCells,
  parseWorksheetXml,
  parseXlsxDataValidations,
  parseXmlPartStrict,
  readPackageRelationships,
  recordDelimitedChange,
  recordDelimitedIssue,
  relationshipsPathFor,
  resolvePackagePartPath,
  scanXlsxNumericTokens,
  summarizePreservedParts,
  validateConditionalFormatRules,
  validateWorkbookPackage,
  validateXlsxNumericValue,
  xlsxExchangeSupportNotice,
} from "./06c-package-and-worksheet-parsing.js";

import {
  strFromU8,
  strToU8,
  zipSync,
  Unzip,
  UnzipInflate,
  unzipSync,
} from "fflate";
import iue from "exceljs";
import lue from "papaparse";
import {
  unwrapDefaultExport,
  sha256HexDigest,
  normalizeDatePolicy,
  dateSystemMaxSerial,
} from "./01-schemas-date-functions-and-validation.js";
import {
  fileImportLimits,
  workbookLimits,
} from "./02-text-and-financial-functions.js";
import {
  assertWorkbookStructureEditable,
  defaultRowHeightOf,
  readHiddenAxisIntervals,
  defaultCalculationSettings,
  validatePrintPageBreaks,
  sortConditionalRulesByPriority,
  normalizeCalculationSettings,
  defaultColumnWidthOf,
  buildNumberFormatCode,
  findOccupiedRangeBounds,
  calculateWorkbookWithSettings,
  requireRetainedXlsxSource,
  encodeBytesToBase64,
  _t,
  aP,
  updateSheetPageSetup,
  cn,
  parseRangeAddress,
  serializeDefinedNames,
  buildDrawingExportNotes,
  applyFilterConditions,
  resolveBorderEdge,
  normalizeWorkbookInput,
  cellBorderEdges,
  columnIndexToLetters,
  buildTableCellEntries,
  worksheetCapacityLimits,
  resolveEffectiveCellStyle,
} from "./03-workbook-model.js";

import {
  inspectionSignatureCache,
  describeImportSource,
  readGeneratedTemplateBuffer,
} from "./07-file-import-and-worker-dispatch.js";

async function exportXltxTemplate(workbook) {
  const r = await exportWorkbookToXlsx(workbook),
    t = convertTemplateToXlsx(
      new Uint8Array(await readGeneratedTemplateBuffer(r.blob)),
      "xlsx",
      "xltx",
    ),
    warnings = [
      ...r.warnings,
      "Derived XLTX template: supported worksheet content only. Macro projects, native automation and unsupported special parts are not transferred. Reopening creates a new workbook; keep native JSON and original source files.",
    ];
  return {
    blob: new Blob([new Uint8Array(t)], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.template",
    }),
    warnings: warnings,
    report: {
      ...r.report,
      format: "xltx",
      fileName: describeImportSource(workbook, "xltx"),
      fileBytes: t.length,
      warnings: warnings,
    },
  };
}
const g7 = "http://schemas.openxmlformats.org/package/2006/relationships",
  contentTypesNamespace =
    "http://schemas.openxmlformats.org/package/2006/content-types",
  escapeXmlAttribute = (e) =>
    e.replace(
      /[&<>"']/g,
      (r) =>
        ({
          "&": "&amp;",
          "<": "&lt;",
          ">": "&gt;",
          '"': "&quot;",
          "'": "&apos;",
        })[r],
    );

function rebuildPreservedPackage(e, r, t) {
  const n = structuredClone(t),
    packageRelationships = readPackageRelationships(e),
    a = readPackageRelationships(r),
    s = { ...r };
  n.phase = "export-result";
  for (const d of ["custom XML", "custom properties"]) {
    const l = n.parts
      .filter(
        (part) =>
          part.family === d ||
          (d === "custom XML" && part.path.startsWith("customXml/")),
      )
      .filter((m) => m.disposition === "preserved-opaque");
    if (!l.length) continue;
    const c = new Set(l.map((m) => m.path)),
      h = packageRelationships.relationships.filter(
        (relationship) =>
          !c.has(relationship.owner) && c.has(relationship.target),
      );
    if (
      l.some((m) => s[m.path] && sha256HexDigest(s[m.path]) !== m.sha256) ||
      h.some((m) =>
        a.relationships.some(
          (relationship) =>
            relationship.owner === m.owner && relationship.type === m.type,
        ),
      )
    ) {
      for (const m of l)
        ((m.disposition = "dropped"),
          (m.reason =
            "Generated package has a conflicting part or relationship. The entire metadata island remains only in the original."));
      continue;
    }
    for (const m of l) s[m.path] = e[m.path];
    for (const m of new Set(h.map((v) => v.owner))) {
      const v = relationshipsPathFor(m),
        g = s[v] ? strFromU8(s[v]) : `<Relationships xmlns="${g7}"/>`,
        b = new Set();
      s[v] &&
        parseXmlPartStrict(s[v], (E) => {
          E.local === "Relationship" && b.add(E.attributes.Id?.value ?? "");
        });
      let y = 1;
      const w = h
        .filter((E) => E.owner === m)
        .map((E) => {
          for (; b.has(`lgPreserve${y}`);) y++;
          const A = `lgPreserve${y++}`;
          return (
            b.add(A),
            `<Relationship xmlns="${g7}" Id="${A}" Type="${escapeXmlAttribute(E.type)}" Target="/${escapeXmlAttribute(E.target)}"/>`
          );
        })
        .join("");
      s[v] = strToU8(expandEmptyXmlElements(g, "Relationships", w));
    }
    const p = l
      .filter((m) => !m.path.endsWith(".rels") && !a.types.has(m.path))
      .map(
        (m) =>
          `<Override xmlns="${contentTypesNamespace}" PartName="/${escapeXmlAttribute(m.path)}" ContentType="${escapeXmlAttribute(m.contentType)}"/>`,
      )
      .join("");
    s["[Content_Types].xml"] = strToU8(
      expandEmptyXmlElements(strFromU8(s["[Content_Types].xml"]), "Types", p),
    );
  }
  return { bytes: zipSync(s), inventory: summarizePreservedParts(n) };
}
function expandEmptyXmlElements(e, r, t) {
  if (!t) return e;
  const n = new RegExp(`</((?:[\\w.-]+:)?${r})\\s*>\\s*$`);
  if (n.test(e)) return e.replace(n, `${t}$&`);
  const o = new RegExp(`<((?:[\\w.-]+:)?${r})([^>]*)/>(\\s*)$`);
  if (o.test(e)) return e.replace(o, `<$1$2>${t}</$1>$3`);
  throw Error("Unable to preserve XLSX package metadata safely.");
}
function describePreservationDisposition(e, r) {
  const t = structuredClone(e);
  t.phase = "export-result";
  for (const part of t.parts)
    part.disposition === "preserved-opaque" &&
      ((part.disposition = "dropped"), (part.reason = r));
  return summarizePreservedParts(t);
}

const countRangeCells = (e) => {
    const rangeAddress = parseRangeAddress(e);
    return (
      (rangeAddress.lastRow - rangeAddress.firstRow + 1) *
      (rangeAddress.lastCol - rangeAddress.firstCol + 1)
    );
  },
  hexColorToArgb = (e) => `FF${e.slice(1).toUpperCase()}`;

function applyValidationDropdownFlags(e, sheet) {
  return e.replace(/<dataValidation\b[^>]*>/g, (t) => {
    const n = /\bsqref="([^"]+)"/.exec(t)?.[1],
      o = sheet.validations?.find(
        (validation) =>
          n &&
          formatRangeCornerCells(validation.range) ===
            formatRangeCornerCells(n),
      );
    return o?.kind === "list" && o.showDropdown === false
      ? t
          .replace(/\s+showDropDown="[^"]*"/, "")
          .replace(/>$/, ' showDropDown="1">')
      : t;
  });
}

function serializeDataValidationRule(e) {
  const r = {
    list: "list",
    number: "decimal",
    integer: "whole",
    textLength: "textLength",
    date: "date",
    time: "time",
    custom: "custom",
  }[e.kind];
  if (
    e.kind === "list" &&
    (e.list.some((t) => /[",\r\n]/.test(t)) || e.list.join(",").length > 253)
  )
    throw Error(
      "This list validation cannot be represented by an inline XLSX list. Retain the native workbook or use shorter values without commas, quotes or line breaks.",
    );
  if (e.kind === "custom" && isExternalReferenceFormula(e.formula))
    throw Error(
      "An unsafe custom validation formula cannot be exported. Keep the native workbook and review the rule.",
    );
  return {
    type: r,
    ...(["list", "custom"].includes(e.kind)
      ? {}
      : { operator: dataValidationOperators[e.operator] }),
    allowBlank: e.allowBlank,
    showErrorMessage: true,
    errorStyle: e.errorStyle ?? "stop",
    errorTitle: e.errorTitle ?? "Invalid value",
    error: e.message,
    showInputMessage: e.showInputMessage !== false,
    promptTitle: e.inputTitle ?? "",
    prompt: e.inputMessage ?? "",
    formulae:
      e.kind === "list"
        ? [JSON.stringify(e.list.join(","))]
        : e.kind === "custom"
          ? [e.formula.slice(1)]
          : [
              e.min,
              ...(["between", "notBetween"].includes(e.operator)
                ? [e.max]
                : []),
            ].map((t) =>
              e.kind === "date"
                ? new Date(Date.UTC(1899, 11, 30) + t * 864e5)
                : t,
            ),
  };
}

function appendStopIfTrueAttribute(e, sheet) {
  const t = sortConditionalRulesByPriority(sheet.conditionalFormats ?? []);
  return e.replace(/<cfRule\b[^>]*>/g, (n) => {
    const o = Number(/\bpriority="(\d+)"/.exec(n)?.[1]),
      a = t[o - 1];
    return a?.stopIfTrue
      ? n.replace(/\s+stopIfTrue="[^"]*"/, "").replace(/>$/, ' stopIfTrue="1">')
      : n;
  });
}
function serializeConditionalFormatRules(e, priority) {
  const t = {
      fill: {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: hexColorToArgb(e.fill) },
      },
      font: { color: { argb: hexColorToArgb(e.color) } },
    },
    rangeAddress = parseRangeAddress(e.range),
    cellAddress = formatCellAddress(
      rangeAddress.firstRow,
      rangeAddress.firstCol,
    ),
    a = (u) => ({
      type: "expression",
      priority: priority,
      formulae: [u.replace(/^=/, "")],
      style: t,
    });
  if (e.kind === "formula") {
    if (isExternalReferenceFormula(e.formula ?? ""))
      throw Error(
        "External-action conditional formulas cannot be exported. Remove that rule or use native JSON.",
      );
    return a(e.formula);
  }
  if (["blank", "notBlank", "error", "notError"].includes(e.kind))
    return a(
      e.kind === "blank"
        ? `LEN(${cellAddress})=0`
        : e.kind === "notBlank"
          ? `LEN(${cellAddress})>0`
          : e.kind === "error"
            ? `ISERROR(${cellAddress})`
            : `NOT(ISERROR(${cellAddress}))`,
    );
  if (e.kind === "unique" || e.kind === "duplicate")
    return a(Wle(e.kind, e.range));
  if (e.kind === "dataBar")
    return {
      type: "dataBar",
      priority: priority,
      cfvo: [{ type: "min" }, { type: "max" }],
      color: { argb: hexColorToArgb(e.fill) },
      gradient: false,
      showValue: true,
      negativeBarColorSameAsPositive: false,
      negativeFillColor: { argb: "FFD13438" },
      axisColor: { argb: "FF777777" },
      axisPosition: "auto",
      direction: "leftToRight",
      minLength: 0,
      maxLength: 100,
    };
  if (e.kind === "iconSet")
    return {
      type: "iconSet",
      priority: priority,
      iconSet: "3Arrows",
      showValue: true,
      reverse: false,
      cfvo: [
        { type: "min" },
        ...(e.iconThresholds ?? [33, 67]).map((u) => ({
          type: e.thresholdType === "number" ? "num" : "percent",
          value: u,
        })),
      ],
    };
  if (e.kind === "contains")
    return a(buildContainsTextFormula(e.value, cellAddress));
  if (e.kind === "colorScale")
    return {
      type: "colorScale",
      priority: priority,
      cfvo: [{ type: "min" }, { type: "max" }],
      color: [
        { argb: hexColorToArgb(e.minColor ?? "#f8e99b") },
        { argb: hexColorToArgb(e.maxColor ?? "#63be7b") },
      ],
    };
  const operator =
      e.kind === "greater"
        ? "greaterThan"
        : e.kind === "less"
          ? "lessThan"
          : e.kind === "equal"
            ? "equal"
            : "between",
    d = (u) =>
      u.trim() !== "" && Number.isFinite(Number(u))
        ? Number(u)
        : `"${u.replace(/"/g, '""')}"`;
  return {
    type: "cellIs",
    operator: operator,
    formulae: [d(e.value), ...(e.kind === "between" ? [d(e.second)] : [])],
    priority: priority,
    style: t,
  };
}
function serializeWorksheetRuleXml(e, sheet) {
  let t = 0,
    n = 0;
  for (const d of sheet.merges ?? []) {
    if ((t += countRangeCells(d)) > workbookLimits.maxCells)
      throw Error("Merged ranges exceed the 100,000-cell interchange budget.");
    const rangeAddress = parseRangeAddress(d);
    for (const [l, cell] of Object.entries(sheet.cells)) {
      const h = _t(l);
      if (
        h.row >= rangeAddress.firstRow &&
        h.row <= rangeAddress.lastRow &&
        h.col >= rangeAddress.firstCol &&
        h.col <= rangeAddress.lastCol &&
        (h.row !== rangeAddress.firstRow || h.col !== rangeAddress.firstCol) &&
        ((cell.raw !== null && cell.raw !== "") ||
          cell.note ||
          cell.tickmark ||
          cell.hyperlink)
      )
        throw Error(
          "Exporting this merged area would discard populated cells or annotations. Unmerge it or preserve those values before XLSX export.",
        );
    }
    e.mergeCells(d);
  }
  const o = [];
  for (const d of sheet.validations ?? []) {
    const rangeAddress = parseRangeAddress(d.range);
    if (
      o.some(
        (l) =>
          l.firstRow <= rangeAddress.lastRow &&
          l.lastRow >= rangeAddress.firstRow &&
          l.firstCol <= rangeAddress.lastCol &&
          l.lastCol >= rangeAddress.firstCol,
      )
    )
      throw Error(
        "Overlapping validation rules cannot be represented by one Excel rule per cell. Use native JSON or resolve overlapping rules before export.",
      );
    if (
      (o.push(rangeAddress),
      (n += countRangeCells(d.range)) > workbookLimits.maxCells)
    )
      throw Error(
        "XLSX validation ranges exceed the 100,000-cell interchange budget. Use native JSON to preserve larger rules.",
      );
    e.dataValidations.add(d.range, serializeDataValidationRule(d));
  }
  const a = new Map(
    sortConditionalRulesByPriority(sheet.conditionalFormats ?? []).map(
      (d, u) => [d.id, u + 1],
    ),
  );
  let s = 0;
  for (const d of sheet.conditionalFormats ?? []) {
    if (d.kind !== "duplicate" && d.kind !== "unique") continue;
    const rangeAddress = parseRangeAddress(d.range);
    for (const [l, cell] of Object.entries(sheet.cells)) {
      if (++s > 1e6)
        throw Error(
          "Duplicate/unique conditional export exceeds 1,000,000 population checks. Use native JSON or fewer rules.",
        );
      const h = _t(l);
      if (
        h.row < rangeAddress.firstRow ||
        h.row > rangeAddress.lastRow ||
        h.col < rangeAddress.firstCol ||
        h.col > rangeAddress.lastCol
      )
        continue;
      let raw = cell.raw;
      if (
        (typeof cell.raw == "string" &&
          cell.raw.startsWith("=") &&
          !cell.literal &&
          (raw = e.getCell(l).result),
        !Hle(raw))
      )
        throw Error(
          `Duplicate/unique conditional rule at ${l} cannot preserve Excel comparison semantics for this value. Use native JSON; XLSX supports finite numbers, canonical numeric text and short printable ASCII text without date-like numeric prefixes.`,
        );
    }
  }
  (sheet.conditionalFormats ?? []).forEach((d) =>
    e.addConditionalFormatting({
      ref: d.range,
      rules: [serializeConditionalFormatRules(d, a.get(d.id))],
    }),
  );
}

function serializeCalculationProperties(e, workbookInput) {
  const calculationSettings = normalizeCalculationSettings(
      workbookInput.calculation ?? defaultCalculationSettings,
    ),
    n = `<calcPr fullPrecision="1" calcId="0" calcMode="${calculationSettings.mode === "manual" ? "manual" : "auto"}" fullCalcOnLoad="1" forceFullCalc="1" calcOnSave="${calculationSettings.mode === "manual" ? "0" : "1"}" iterate="${calculationSettings.iterative ? "1" : "0"}" iterateCount="${calculationSettings.maxIterations}" iterateDelta="${calculationSettings.maxChange}"/>`;
  return /<(?:\w+:)?calcPr\b[^>]*\/>/.test(e)
    ? e.replace(/<(?:\w+:)?calcPr\b[^>]*\/>/, n)
    : /<(?:\w+:)?calcPr\b[^>]*>[\s\S]*?<\/(?:\w+:)?calcPr>/.test(e)
      ? e.replace(/<(?:\w+:)?calcPr\b[^>]*>[\s\S]*?<\/(?:\w+:)?calcPr>/, n)
      : e.replace(/<\/(?:\w+:)?workbook>/, `${n}$&`);
}
function verifyAutomationInspection(inspection) {
  const r = inspectionSignatureCache.get(inspection);
  if (!r || r.signature !== JSON.stringify(inspection))
    throw Error(
      "Automation inspection was modified or is untrusted. Inspect the original bytes again.",
    );
  return r.bytes.slice();
}
function retainSourceArtifact(inspection, actor) {
  const t = verifyAutomationInspection(inspection);
  if (!inspection.retainedEligible)
    throw Error(
      "This source exceeds the 2 MB artifact-retention limit. Keep the original file separately.",
    );
  if (!actor.trim() || actor.length > 100)
    throw Error("A bounded reviewer name is required.");
  return {
    id: crypto.randomUUID(),
    fileName: inspection.fileName,
    kind: inspection.kind,
    sha256: inspection.sha256,
    byteLength: t.length,
    base64: encodeBytesToBase64(t),
    retainedAt: new Date().toISOString(),
    retainedBy: actor,
  };
}
function inspectEmbeddedObjects(inspection) {
  verifyAutomationInspection(inspection);
  const r = inspection.items.filter((item) =>
    /embedded|activeX|OLE|External automation/i.test(item.kind),
  );
  if (!r.length)
    throw Error(
      "No known OLE, ActiveX or embedded-object references were detected. Separately stored objects remain unknown.",
    );
  return {
    id: crypto.randomUUID(),
    fileName: inspection.fileName,
    fileBytes: inspection.fileBytes,
    sourceSha256: inspection.sha256,
    inspectedAt: new Date().toISOString(),
    execution: "never",
    items: r.map((t) => ({ ...t, execution: "never" })),
    warnings: [
      ...inspection.warnings,
      "Inventory only: no payload is opened, fetched or executed. Missing parts and external references remain unresolved. Keep the original source separately.",
    ],
  };
}
function retainEmbeddedObjectInventory(workbook, inspection, t = {}) {
  if (t.readOnly) throw Error("This workbook is read-only.");
  assertWorkbookStructureEditable(workbook);
  const n = inspectEmbeddedObjects(inspection);
  if (
    workbook.embeddedObjectInventories?.some(
      (embeddedObjectInventory) =>
        embeddedObjectInventory.sourceSha256 === n.sourceSha256,
    )
  )
    throw Error("This source inventory has already been retained.");
  return normalizeWorkbookInput({
    ...workbook,
    embeddedObjectInventories: [
      ...(workbook.embeddedObjectInventories ?? []),
      n,
    ],
  });
}

const excelErrorLiterals = new Set([
  "#NULL!",
  "#DIV/0!",
  "#VALUE!",
  "#REF!",
  "#NAME?",
  "#NUM!",
  "#N/A",
]);
function toExportedCellValue(e) {
  if (e.type !== "blank")
    return e.type === "error"
      ? excelErrorLiterals.has(String(e.value))
        ? { error: e.value }
        : undefined
      : e.value;
}
function classifyArrayFormula(e, r) {
  const t = (o, a) => {
    const s = new RegExp(`(?:^|\\s)${a}\\s*=\\s*(["'])(.*?)\\1`, "i").exec(
      o,
    )?.[2];
    return s?.replace(/&#(x[0-9a-f]+|[0-9]+);/gi, (u, l) =>
      String.fromCodePoint(
        l[0].toLowerCase() === "x" ? parseInt(l.slice(1), 16) : Number(l),
      ),
    );
  };
  let n;
  for (const o of e.matchAll(
    /<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>|<(\/?)(?:\w+:)?(c|f)\b([^>]*)>/gi,
  )) {
    if (!o[2]) continue;
    if (o[2].toLowerCase() === "c") {
      n = o[1] || /\/\s*$/.test(o[3]) ? undefined : o[3];
      continue;
    }
    if (n === undefined || o[1] || t(o[3], "t") !== "array") continue;
    const a = t(n, "r") ?? "unknown address",
      s = t(o[3], "ref") ?? a,
      d =
        t(n, "cm") !== undefined
          ? "Metadata-qualified Excel array formula (dynamic/legacy classification requires Excel metadata)"
          : "Legacy CSE array formula";
    throw new Error(
      `${d} at ${r}!${a}, range ${s}. Import stopped before changing the workbook. Open the original in Excel and convert the array to ordinary supported formulas or an explicitly reviewed values snapshot, then import that copy. Fixed CSE ranges and Excel array metadata are not silently treated as dynamic spills.`,
    );
  }
}

const oue = unwrapDefaultExport(iue);

const x7 = unwrapDefaultExport(lue);

function dateSerialToIsoText(raw, dateOrder, t = "1900") {
  if (!Number.isFinite(raw) || raw < 0 || raw >= dateSystemMaxSerial(t) + 1)
    throw Error(
      `Date serial is outside 0–${dateSystemMaxSerial(t)} for the ${t} system.`,
    );
  let n = Math.floor(raw),
    o = Math.round((raw - n) * 86400);
  if ((o === 86400 && (n++, (o = 0)), n > dateSystemMaxSerial(t)))
    throw Error("Rounded date/time exceeds year 9999.");
  const a =
      t === "1900" && n === 60
        ? "1900-02-29"
        : new Date(
            t === "1904"
              ? Date.UTC(1904, 0, 1) + n * 864e5
              : Date.UTC(1899, 11, 31) + (n > 60 ? n - 1 : n) * 864e5,
          )
            .toISOString()
            .slice(0, 10),
    [s, d, u] = a.split("-");
  return (
    (dateOrder === "ymd"
      ? a
      : dateOrder === "dmy"
        ? `${u}/${d}/${s}`
        : `${d}/${u}/${s}`) +
    (o
      ? ` ${String(Math.floor(o / 3600)).padStart(2, "0")}:${String(Math.floor(o / 60) % 60).padStart(2, "0")}:${String(o % 60).padStart(2, "0")}`
      : "")
  );
}

function getWorksheetUsedRange(workbook, sheetId) {
  const sheet = workbook.sheets.find((sheet) => sheet.id === sheetId);
  if (!sheet) throw Error("Worksheet not found.");
  const occupiedRangeBounds = findOccupiedRangeBounds(sheet, false);
  return `A1:${columnIndexToLetters(Math.max(occupiedRangeBounds?.lastCol ?? 0, (sheet.delimitedSource?.columns ?? 1) - 1))}${Math.max(occupiedRangeBounds?.lastRow ?? 0, (sheet.delimitedSource?.records ?? 1) - 1) + 1}`;
}
function encodeExportTextBytes(e, encoding, bom) {
  if (
    !["utf-8", "utf-16le", "utf-16be", "windows-1252"].includes(encoding) ||
    typeof bom != "boolean"
  )
    throw Error("Invalid output encoding or BOM setting.");
  for (const s of e)
    if (s.length === 1 && s.charCodeAt(0) >= 55296 && s.charCodeAt(0) <= 57343)
      throw Error(
        "Unpaired Unicode surrogate cannot be exported; correct the text first.",
      );
  if (encoding === "utf-8")
    return new TextEncoder().encode((bom ? "\uFEFF" : "") + e);
  if (encoding === "windows-1252") {
    if (bom)
      throw Error(
        "Windows-1252 does not support a BOM; disable BOM or choose UTF-8.",
      );
    const s = new Map(byteCharacterTable.map((u, l) => [u, l])),
      d = [];
    for (const u of e) {
      const l = s.get(u);
      if (l === undefined)
        throw Error(
          "A character cannot be represented in Windows-1252. Choose UTF-8; text is never replaced.",
        );
      d.push(l);
    }
    return Uint8Array.from(d);
  }
  const n = (bom ? "\uFEFF" : "") + e,
    o = new Uint8Array(n.length * 2),
    a = new DataView(o.buffer);
  for (let s = 0; s < n.length; s++)
    a.setUint16(s * 2, n.charCodeAt(s), encoding === "utf-16le");
  return o;
}
const needsFormulaInjectionGuard = (e) =>
  /^[\s\u0000-\u001f]*[=+\-@＝＋－＠]/u.test(e) || /^[\t\r\n]/.test(e);
function quoteDelimitedField(e, delimitedOptions) {
  const t =
    e.includes(delimitedOptions.delimiter) ||
    /[\r\n\uFEFF]/.test(e) ||
    (delimitedOptions.quote !== "none" && e.includes(delimitedOptions.quote)) ||
    /^[ ]|[ ]$/.test(e) ||
    delimitedOptions.quoteAll;
  if (delimitedOptions.quote === "none") {
    if (e.includes(delimitedOptions.delimiter) || /[\r\n]/.test(e))
      throw Error(
        "A field needs quoting; choose a quote character to preserve delimiters/newlines.",
      );
    return e;
  }
  if (!t) return e;
  const n =
    delimitedOptions.escape === "double"
      ? e
          .split(delimitedOptions.quote)
          .join(delimitedOptions.quote + delimitedOptions.quote)
      : e
          .replaceAll("\\", "\\\\")
          .split(delimitedOptions.quote)
          .join("\\" + delimitedOptions.quote);
  return delimitedOptions.quote + n + delimitedOptions.quote;
}
function exportWorksheetToCsv(workbook, sheetId, t = {}) {
  const workbookInput = normalizeWorkbookInput(workbook),
    sheet = workbookInput.sheets.find((sheet) => sheet.id === sheetId);
  if (!sheet) throw Error("Worksheet not found.");
  const delimitedOptions = {
    ...normalizeDelimitedOptions({
      ...t,
      delimiter: t.delimiter ?? ",",
      encoding: t.encoding ?? "utf-8",
      newline:
        t.newline ??
        `\r
`,
    }),
    includeHeader: t.includeHeader ?? true,
    bom: t.bom ?? false,
    quoteAll: t.quoteAll ?? false,
    dateOutput: t.dateOutput ?? "calendar",
    range:
      t.range ??
      t.population?.definition.range ??
      getWorksheetUsedRange(workbookInput, sheetId),
    ...(t.population ? { population: t.population } : {}),
  };
  for (const w of ["includeHeader", "bom", "quoteAll"])
    if (typeof delimitedOptions[w] != "boolean")
      throw Error(`Invalid export ${w} control.`);
  if (!["calendar", "serial"].includes(delimitedOptions.dateOutput))
    throw Error("Invalid date export mode.");
  (delimitedOptions.delimiter === "auto" && (delimitedOptions.delimiter = ","),
    delimitedOptions.encoding === "auto" &&
      (delimitedOptions.encoding = "utf-8"),
    delimitedOptions.newline === "auto" &&
      (delimitedOptions.newline = `\r
`));
  const rangeAddress = parseRangeAddress(delimitedOptions.range),
    d = rangeAddress.lastRow - rangeAddress.firstRow + 1,
    u = rangeAddress.lastCol - rangeAddress.firstCol + 1;
  if (d * u > workbookLimits.maxCells)
    throw Error(
      "Delimited/CSV export would expand more than 100,000 cells including blank gaps. Choose a smaller range or use native JSON.",
    );
  const delimitedImportReport = buildDelimitedImportReport(
      "export",
      `${sheet.name}.${delimitedOptions.delimiter === "	" ? "tsv" : "csv"}`,
      0,
      delimitedOptions,
    ),
    workbookWithSettings = calculateWorkbookWithSettings(workbookInput),
    h = [];
  if (delimitedOptions.population) {
    if (
      !["full", "visible"].includes(delimitedOptions.population.scope) ||
      !delimitedOptions.hasHeader
    )
      throw Error(
        "Filtered exports require a header and an explicit full or visible scope.",
      );
    const context = applyFilterConditions(
      workbookInput,
      sheetId,
      delimitedOptions.population.definition,
      workbookWithSettings,
    ).context;
    if (
      delimitedOptions.range !== context.range &&
      delimitedOptions.range !== delimitedOptions.population.definition.range
    )
      throw Error(
        "The export range must match the declared filter population.",
      );
    ((delimitedImportReport.population = {
      ...context,
      scope: delimitedOptions.population.scope,
      exportedSourceRows:
        delimitedOptions.population.scope === "visible"
          ? context.visibleSourceRows
          : context.sourceRows,
    }),
      delimitedImportReport.warnings.push(
        `Population scope: ${delimitedOptions.population.scope}; exporting ${delimitedImportReport.population.exportedSourceRows.length} of ${context.totalRows} source data rows. Save the schema report for filter criteria, row locators and source IDs. Session filter criteria are not encoded in CSV/TSV.`,
      ));
  } else
    delimitedImportReport.warnings.push(
      "Population scope: full selected rectangle, including hidden and filtered rows.",
    );
  const f = delimitedImportReport.population
    ? new Set(delimitedImportReport.population.exportedSourceRows)
    : null;
  ((delimitedImportReport.columns = u),
    delimitedImportReport.warnings.push(
      "Values-only export: formulas are evaluated; formatting, notes, history and workbook metadata are not carried by CSV/TSV. Reimport with the visible schema; the file itself has no type schema.",
      "Safe export always prefixes an apostrophe to formula-like text, including whitespace/control-prefixed and full-width triggers. Numeric negatives stay numeric. Quotes alone do not neutralize spreadsheet formulas. Changed text is listed below; it is not automatically unescaped on reimport.",
    ),
    sheet.delimitedSource &&
      delimitedImportReport.warnings.push(
        "The default range includes the historical imported footprint so trailing blank records survive. Review the range after structural edits; this is not a live table binding.",
      ));
  let p = 0;
  for (
    let firstRow = rangeAddress.firstRow;
    firstRow <= rangeAddress.lastRow;
    firstRow++
  ) {
    if (firstRow !== rangeAddress.firstRow && f && !f.has(firstRow + 1))
      continue;
    const record = firstRow - rangeAddress.firstRow + 1,
      A = delimitedOptions.hasHeader && firstRow === rangeAddress.firstRow,
      k = [],
      raw = [],
      values = [];
    let O = true;
    for (let V = 0; V < u; V++) {
      const D = `${columnIndexToLetters(rangeAddress.firstCol + V)}${firstRow + 1}`,
        B = sheet.cells[D],
        _ = workbookWithSettings(sheetId, D),
        C = delimitedOptions.columns[V] ?? "auto";
      if (
        (firstRow === rangeAddress.firstRow &&
          delimitedImportReport.schema.push({
            index: V,
            label: delimitedOptions.hasHeader
              ? String(_ ?? "")
              : columnIndexToLetters(rangeAddress.firstCol + V),
            type: C,
            observed: [],
          }),
        A && !delimitedOptions.includeHeader)
      )
        continue;
      let T = String(_ ?? ""),
        N = typeof _ == "string";
      try {
        const cell =
          A || C === "auto"
            ? { raw: _ }
            : typeof _ == "number" && ["number", "integer", "date"].includes(C)
              ? { raw: _ }
              : parseDelimitedCellValue(
                  T,
                  C,
                  delimitedOptions,
                  workbookInput.datePolicy?.dateSystem,
                );
        if (
          !A &&
          C === "integer" &&
          typeof cell.raw == "number" &&
          !Number.isSafeInteger(cell.raw)
        )
          throw Error("Integer column requires a safe whole number.");
        if (
          ((N = typeof cell.raw == "string"),
          (T = String(cell.raw ?? "")),
          typeof cell.raw == "number" &&
            (!A &&
            (C === "date" || (C === "auto" && B?.format === "date")) &&
            delimitedOptions.dateOutput === "calendar"
              ? (T = dateSerialToIsoText(
                  cell.raw,
                  delimitedOptions.dateOrder,
                  workbookInput.datePolicy?.dateSystem,
                ))
              : (T = T.replace(".", delimitedOptions.decimal))),
          N && needsFormulaInjectionGuard(T))
        ) {
          const $ = T;
          ((T = "'" + T),
            delimitedImportReport.safeTextChanges++,
            recordDelimitedChange(
              delimitedImportReport,
              record,
              V + 1,
              D,
              "Formula-like text prefixed with apostrophe",
              $,
              T,
            ));
        } else
          T !== String(_ ?? "") &&
            recordDelimitedChange(
              delimitedImportReport,
              record,
              V + 1,
              D,
              "Selected type, number or date representation",
              String(_ ?? ""),
              T,
            );
        const L = cell.raw === null ? "blank" : typeof cell.raw;
        if (
          (delimitedImportReport.schema[V].observed.includes(L) ||
            delimitedImportReport.schema[V].observed.push(L),
          T.length > fileImportLimits.maxCellTextLength)
        )
          throw Error(
            "Safe-export prefix would exceed the 32,767-character field limit. Shorten the text before exporting.",
          );
        (T !== "" && (O = false),
          k.push(quoteDelimitedField(T, delimitedOptions)));
      } catch (error) {
        (recordDelimitedIssue(
          delimitedImportReport,
          record,
          V + 1,
          error instanceof Error ? error.message : "Export conversion failed.",
          T,
        ),
          k.push(""));
      }
      (_ !== null && delimitedImportReport.storedCells++,
        V < 200 && (raw.push(String(_ ?? "")), values.push(T)));
    }
    if (A && !delimitedOptions.includeHeader) continue;
    const j = k.join(delimitedOptions.delimiter);
    O && delimitedImportReport.blankRecords++;
    const z =
      j === "" && delimitedOptions.quote !== "none"
        ? delimitedOptions.quote + delimitedOptions.quote
        : j;
    if (
      ((p +=
        new TextEncoder().encode(z).length +
        (h.length ? delimitedOptions.newline.length : 0)),
      p > workbookLimits.maxFileBytes * 3)
    )
      throw Error(
        "Delimited export exceeds the 10 MB output limit. Choose a smaller range.",
      );
    (h.push(z),
      delimitedImportReport.records++,
      delimitedImportReport.preview.length < 25 &&
        delimitedImportReport.preview.push({
          record: record,
          raw: raw,
          values: values,
        }));
  }
  let m = h.join(delimitedOptions.newline);
  h.at(-1) === "" &&
    delimitedOptions.quote === "none" &&
    (m += delimitedOptions.newline);
  let v = new Uint8Array();
  try {
    v = encodeExportTextBytes(
      m,
      delimitedOptions.encoding,
      delimitedOptions.bom,
    );
  } catch (error) {
    recordDelimitedIssue(
      delimitedImportReport,
      0,
      undefined,
      error instanceof Error ? error.message : "Encoding failure.",
    );
  }
  if (v.length > workbookLimits.maxFileBytes)
    throw Error(
      "CSV/TSV export exceeds the 10 MB output limit. Choose a smaller range or native JSON.",
    );
  return (
    (delimitedImportReport.fileBytes = v.length),
    (delimitedImportReport.valid = delimitedImportReport.issueCount === 0),
    (delimitedImportReport.complete = true),
    {
      text: m,
      bytes: delimitedImportReport.valid ? v : new Uint8Array(),
      report: delimitedImportReport,
    }
  );
}

function importXlsxWorkbook(bytes, r) {
  const t = inspectXlsxPackage(bytes),
    n = new Set(t.keys()),
    size = t.size,
    workbookParseState = createWorkbookParseState();
  if (
    [...n].some((w) =>
      /(?:vbaProject\.bin|xl\/macrosheets\/|xl\/activeX\/|xl\/embeddings\/)/i.test(
        w,
      ),
    )
  )
    throw new Error(
      "Workbooks containing macros, ActiveX or embedded objects are not accepted. Save a clean XLSX copy first.",
    );
  let s = 0,
    d,
    u = 0,
    l = 0,
    c = 0,
    h = 0,
    f = 0,
    p = 0,
    m = 0,
    v = false;
  const g = new Set(),
    b = new Unzip((w) => {
      if ((u++, !n.has(w.name) || g.has(w.name))) {
        d = new Error("Archive entries do not match its directory.");
        return;
      }
      g.add(w.name);
      let E = 0;
      const A = /\.(?:xml|rels|vml)$/i.test(w.name),
        k = A ? new TextDecoder("utf-8", { fatal: true }) : undefined;
      let F = -1,
        I = "";
      ((w.ondata = (O, j, z) => {
        if (O) {
          d = new Error(
            `Corrupt XLSX compressed content in ${w.name}. Repair the original workbook and save a clean XLSX copy.`,
          );
          return;
        }
        if (((s += j.length), (E += j.length), s > maxXlsxExpandedByteLimit)) {
          (w.terminate(),
            (d = new Error("The XLSX archive expands beyond 50 MB.")));
          return;
        }
        if (E > t.get(w.name).size) {
          (w.terminate(),
            (d = new Error(
              "The XLSX archive expansion does not match its directory.",
            )));
          return;
        }
        F = computeCrc32(F, j);
        try {
          k && (I += k.decode(j, { stream: !z }));
        } catch {
          d = new Error(
            `Unsupported XML encoding in ${w.name}. Save a standard UTF-8 XLSX copy.`,
          );
          return;
        }
        if (z) {
          if (E !== t.get(w.name).size) {
            d = new Error(
              "The XLSX archive expansion does not match its directory.",
            );
            return;
          }
          if ((F ^ -1) >>> 0 !== t.get(w.name).crc) {
            d = new Error(
              `XLSX checksum mismatch in ${w.name}. The archive is corrupt; repair the original and save a clean XLSX copy.`,
            );
            return;
          }
          try {
            if (
              (A && parseWorksheetXml(I, w.name, t, workbookParseState),
              /^xl\/styles\.xml$/i.test(w.name))
            )
              for (const D of I.matchAll(/<(?:\w+:)?sz\b([^>]*)>/g)) {
                const B = Number(getXmlAttribute(D[1], "val"));
                (!Number.isInteger(B) || B < 8 || B > 72) && (v = true);
              }
            if (/^xl\/worksheets\//i.test(w.name)) {
              (classifyArrayFormula(I, w.name),
                (m = parseWorksheetCells(I, w.name, m)),
                /^xl\/worksheets\/[^/]+\.xml$/i.test(w.name) &&
                  r.numericTokens.set(w.name, scanXlsxNumericTokens(I)));
              for (const C of I.matchAll(/<(?:\w+:)?pageSetup\b([^>]*)>/gi)) {
                const xmlAttribute = getXmlAttribute(C[1], "pageOrder");
                if (
                  xmlAttribute !== undefined &&
                  !["overThenDown", "downThenOver"].includes(
                    decodeXmlEntities(xmlAttribute),
                  )
                )
                  throw new Error(
                    "Unsupported XLSX page order. Import stopped to preserve the source layout.",
                  );
              }
              const D = {};
              for (const [C, T, N, R] of [
                [
                  "rowBreaks",
                  "rowBreaks",
                  workbookLimits.maxRows,
                  workbookLimits.maxCols,
                ],
                [
                  "colBreaks",
                  "columnBreaks",
                  workbookLimits.maxCols,
                  workbookLimits.maxRows,
                ],
              ]) {
                const L = [
                  ...I.matchAll(
                    new RegExp(
                      "<(?:\\w+:)?" +
                        C +
                        "\\b[^>]*>([\\s\\S]*?)</(?:\\w+:)?" +
                        C +
                        ">",
                      "gi",
                    ),
                  ),
                ];
                if (L.length > 1)
                  throw Error(
                    "Duplicate XLSX page-break sections are not supported.",
                  );
                const $ = [];
                let U = 0;
                for (const H of (L[0]?.[1] ?? "").matchAll(
                  /<(?:\w+:)?brk\b([^>]*)\/?\s*>/gi,
                )) {
                  if (++U > 100)
                    throw Error(
                      "XLSX page breaks exceed 100 boundaries per axis.",
                    );
                  const xmlAttribute = getXmlAttribute(H[1], "id"),
                    X = getXmlAttribute(H[1], "man");
                  if (
                    !/^[1-9]\d*$/.test(xmlAttribute ?? "") ||
                    !Number.isSafeInteger(Number(xmlAttribute)) ||
                    Number(xmlAttribute) >= N
                  )
                    throw Error("Invalid XLSX page-break boundary.");
                  if (X !== "1" && X !== "true") continue;
                  const J = Number(getXmlAttribute(H[1], "min") ?? 0),
                    q = Number(getXmlAttribute(H[1], "max") ?? R - 1);
                  if (J !== 0 || !Number.isSafeInteger(q) || q < R - 1)
                    throw Error(
                      "Partial-span XLSX page breaks are not supported. Use full-row or full-column breaks.",
                    );
                  $.push(Number(xmlAttribute));
                }
                $.sort((H, Q) => H - Q);
                const G = validatePrintPageBreaks($, N);
                G && (D[T] = G);
              }
              (r.printBreaks.set(w.name, D),
                r.conditionalStops.set(
                  w.name,
                  validateConditionalFormatRules(I),
                ),
                r.validationAreas.set(w.name, parseXlsxDataValidations(I)));
              const B = [],
                _ = [];
              for (const C of I.matchAll(/<(?:\w+:)?col\b([^>]*)>/gi)) {
                const xmlAttribute = getXmlAttribute(C[1], "min"),
                  N = getXmlAttribute(C[1], "max"),
                  R = getXmlAttribute(C[1], "width"),
                  start = Number(xmlAttribute) - 1,
                  end = Number(N) - 1;
                if (
                  !/^[1-9]\d*$/.test(xmlAttribute ?? "") ||
                  !/^[1-9]\d*$/.test(N ?? "") ||
                  !Number.isSafeInteger(start) ||
                  !Number.isSafeInteger(end) ||
                  start < 0 ||
                  end < start ||
                  end >= workbookLimits.maxCols ||
                  _.length >= workbookLimits.maxCols
                )
                  throw Error(
                    "XLSX column metadata must stay within 16,384 nonoverlapping columns.",
                  );
                const width = R === undefined ? undefined : Number(R);
                if (
                  width !== undefined &&
                  (!R?.trim() || !Number.isFinite(width) || width < 0)
                )
                  throw Error(
                    "XLSX column widths must be finite nonnegative values.",
                  );
                _.push({
                  start: start,
                  end: end,
                  ...(width === undefined ? {} : { width: width }),
                });
              }
              _.sort((C, T) => C.start - T.start);
              for (let C = 1; C < _.length; C++)
                if (_[C].start <= _[C - 1].end)
                  throw Error(
                    "XLSX column metadata must stay within 16,384 nonoverlapping columns.",
                  );
              if (
                (r.columns.set(w.name, _),
                /<(?:\w+:)?sheetFormatPr\b[^>]*\bzeroHeight\s*=\s*["'](?:1|true)["']/i.test(
                  I,
                ))
              )
                throw new Error(
                  "Default-hidden worksheet rows cannot be imported safely. Replace the default with explicit hidden rows before importing.",
                );
              for (const C of I.matchAll(/<(?:\w+:)?row\b([^>]*)>/gi)) {
                if (!/\bhidden\s*=\s*["'](?:1|true)["']/i.test(C[1])) continue;
                const T = Number(getXmlAttribute(C[1], "r")) - 1;
                if (
                  !Number.isInteger(T) ||
                  T < 0 ||
                  T >= workbookLimits.maxRows
                )
                  throw new Error(
                    "Hidden XLSX rows must have valid row numbers.",
                  );
                const N = B.at(-1);
                if (N && T <= N.end)
                  throw new Error(
                    "Hidden XLSX row numbers must be distinct and ascending.",
                  );
                if (
                  (N && T === N.end + 1
                    ? (N.end = T)
                    : B.push({ start: T, end: T }),
                  B.length > worksheetCapacityLimits.maxIntervalsPerAxis)
                )
                  throw new Error(
                    "Hidden metadata exceeds 10,000 intervals per worksheet axis.",
                  );
                if (++p > worksheetCapacityLimits.maxXlsxHiddenRows)
                  throw new Error(
                    "XLSX hidden rows exceed the 100,000-row processing budget. Reduce hidden rows in the source; native JSON preserves larger visibility intervals.",
                  );
              }
              B.length && r.hiddenRows.set(w.name, B);
            }
            if (w.name === "xl/workbook.xml")
              for (const D of I.matchAll(/<(?:\w+:)?sheet\b([^>]*)>/g)) {
                const xmlAttribute = getXmlAttribute(D[1], "name"),
                  _ = getXmlAttribute(D[1], "r:id");
                xmlAttribute &&
                  _ &&
                  r.sheetNames.set(
                    decodeXmlEntities(xmlAttribute),
                    decodeXmlEntities(_),
                  );
              }
            if (w.name === "xl/_rels/workbook.xml.rels")
              for (const D of I.matchAll(/<(?:\w+:)?Relationship\b([^>]*)>/g)) {
                const xmlAttribute = getXmlAttribute(D[1], "Id"),
                  _ = getXmlAttribute(D[1], "Target");
                xmlAttribute &&
                  _ &&
                  getXmlAttribute(D[1], "TargetMode") !== "External" &&
                  r.sheetRelationships.set(
                    decodeXmlEntities(xmlAttribute),
                    resolvePackagePartPath(decodeXmlEntities(_)),
                  );
              }
            for (const D of I.matchAll(
              /<(?:\w+:)?(mergeCell|dataValidation)\s[^>]*?\b(ref|sqref)\s*=\s*["']([^"']+)["'][^>]*>/g,
            ))
              for (const B of decodeXmlEntities(D[3]).trim().split(/\s+/)) {
                const rangeAddress = parseRangeAddress(B),
                  C =
                    (rangeAddress.lastRow - rangeAddress.firstRow + 1) *
                    (rangeAddress.lastCol - rangeAddress.firstCol + 1);
                if (D[1] === "mergeCell") {
                  if (
                    C > 500 ||
                    ++f > 1e3 ||
                    (c += C) > workbookLimits.maxCells
                  )
                    throw new Error(
                      "Merged ranges exceed the supported import budget (500 cells per merge and 100,000 cells overall).",
                    );
                } else if ((h += C) > workbookLimits.maxCells)
                  throw new Error(
                    "XLSX validation ranges exceed the 100,000-cell import budget. Narrow the validation area in the source workbook.",
                  );
              }
            if (w.name === "xl/workbook.xml") {
              r.calculation = parseCalculationProperties(I);
              const D = aP(I);
              ((r.names = D.names),
                (r.nameWarnings = D.warnings),
                (r.builtinNameXml = D.builtinXml));
            }
          } catch (error) {
            d =
              error instanceof Error
                ? error
                : new Error("Invalid XLSX worksheet metadata.");
            return;
          }
          l++;
        }
      }),
        w.start());
    });
  b.register(UnzipInflate);
  for (let w = 0; w < bytes.length; w += 2048) {
    if (d) throw d;
    try {
      b.push(
        bytes.subarray(w, Math.min(w + 2048, bytes.length)),
        w + 2048 >= bytes.length,
      );
    } catch {
      throw new Error(
        "Corrupt XLSX compressed content. Repair the original workbook and save a clean XLSX copy.",
      );
    }
  }
  if (d) throw d;
  if (u !== size || l !== size)
    throw new Error("The XLSX archive contains incomplete entries.");
  const y = [
    xlsxExchangeSupportNotice,
    ...r.nameWarnings,
    ...validateWorkbookPackage(workbookParseState, n),
  ];
  return (
    (r.emptyStringCaches = workbookParseState.emptyStringCaches),
    v &&
      y.push(
        "Unsupported fractional or out-of-range font sizes were found in the source. The XLSX reader may normalize them; the editor supports integer font sizes from 8 to 72 points.",
      ),
    [...n].some((w) => /externalLinks|connections\.xml/i.test(w)) &&
      y.push(
        "External workbook links or connections were detected. They will not be refreshed or included in the derived copy; the original retains their source parts.",
      ),
    y
  );
}
const expandHexColorToArgb = (e) => {
    const r = e.slice(1);
    return `FF${r.length === 3 ? [...r].map((t) => t + t).join("") : r}`;
  },
  builtinNumberFormatCodes = {
    general: "General",
    number: "#,##0.00",
    currency: "$#,##0.00;[Red]($#,##0.00)",
    percent: "0.00%",
    date: "yyyy-mm-dd",
  };
function serializeWorksheetXml(e, sheets) {
  if (
    !sheets.some(
      (sheet) => sheet.hiddenRows?.length || sheet.hiddenCols?.length,
    )
  )
    return e;
  const t = unzipSync(e);
  return (
    sheets.forEach((sheet, o) => {
      const hiddenAxisIntervals = readHiddenAxisIntervals(sheet, "row");
      if (!hiddenAxisIntervals.length && !sheet.hiddenCols?.length) return;
      const s = "xl/worksheets/sheet" + (o + 1) + ".xml";
      let d = strFromU8(t[s]);
      if (
        ((d = d.replace(/<col\b([^>]*)\/>/g, (p, m) => {
          const v = Number(getXmlAttribute(m, "min")) - 1,
            g = Number(getXmlAttribute(m, "max")) - 1;
          if (
            !Number.isInteger(v) ||
            !Number.isInteger(g) ||
            v < 0 ||
            g < v ||
            g >= workbookLimits.maxCols
          )
            throw Error("The generated worksheet has invalid column metadata.");
          const b = m
              .replace(
                /\s+(?:min|max|width|customWidth)\s*=\s*(["'])[\s\S]*?\1/g,
                "",
              )
              .replace(/\/\s*$/, ""),
            y = [];
          for (let A = v; A <= g;) {
            const k = sheet.columnWidths?.[A];
            let F = A;
            for (; F < g && sheet.columnWidths?.[F + 1] === k;) F++;
            (y.push(
              "<col" +
                b +
                ' min="' +
                (A + 1) +
                '" max="' +
                (F + 1) +
                '"' +
                (k === undefined
                  ? ' customWidth="0"'
                  : ' width="' + k / 7 + '" customWidth="1"') +
                "/>",
            ),
              (A = F + 1));
          }
          return y.join("");
        })),
        !hiddenAxisIntervals.length)
      ) {
        t[s] = strToU8(d);
        return;
      }
      const u =
        /<sheetData\b[^>]*\/>|<sheetData\b[^>]*>([\s\S]*?)<\/sheetData>/.exec(
          d,
        );
      if (!u)
        throw new Error("The generated worksheet is missing its data section.");
      const l = new Map();
      for (const p of (u[1] ?? "").matchAll(
        /<row\b[^>]*(?:\/>|>[\s\S]*?<\/row>)/g,
      )) {
        const m = Number(
          getXmlAttribute(/^<row\b([^>]*)>/.exec(p[0])?.[1] ?? "", "r"),
        );
        if (!Number.isInteger(m) || m < 1 || m > workbookLimits.maxRows)
          throw new Error("The generated worksheet has an invalid row number.");
        l.set(m, p[0]);
      }
      for (const hiddenAxisInterval of hiddenAxisIntervals)
        for (
          let start = hiddenAxisInterval.start;
          start <= hiddenAxisInterval.end;
          start++
        ) {
          const v = start + 1,
            g = l.get(v);
          l.set(
            v,
            g
              ? g.replace(
                  /^<row\b([^>]*)>/,
                  (b, y) =>
                    "<row" +
                    y
                      .replace(/\s+hidden\s*=\s*(["'])[\s\S]*?\1/g, "")
                      .replace(/\/\s*$/, "") +
                    ' hidden="1"' +
                    (/\/\s*$/.test(y) ? "/>" : ">"),
                )
              : '<row r="' + v + '" hidden="1"/>',
          );
        }
      const c = [...l]
        .sort(([p], [m]) => p - m)
        .map(([, p]) => p)
        .join("");
      t[s] = strToU8(d.replace(u[0], "<sheetData>" + c + "</sheetData>"));
    }),
    zipSync(t)
  );
}
function applyPageBreakXml(e, sheets) {
  if (
    !sheets.some(
      (sheet) =>
        sheet.pageLayout?.rowBreaks?.length ||
        sheet.pageLayout?.columnBreaks?.length,
    )
  )
    return e;
  const t = unzipSync(e);
  return (
    sheets.forEach((sheet, o) => {
      const pageLayout = sheet.pageLayout;
      if (!pageLayout?.rowBreaks?.length && !pageLayout?.columnBreaks?.length)
        return;
      const s = `xl/worksheets/sheet${o + 1}.xml`;
      let d = strFromU8(t[s]);
      d = d.replace(
        /<(?:rowBreaks|colBreaks)\b[^>]*(?:\/>|>[\s\S]*?<\/(?:rowBreaks|colBreaks)>)/g,
        "",
      );
      const l = [
        ["rowBreaks", pageLayout.rowBreaks, 16383],
        ["colBreaks", pageLayout.columnBreaks, 1048575],
      ]
        .map(([f, p, m]) =>
          p?.length
            ? `<${f} count="${p.length}" manualBreakCount="${p.length}">${p.map((v) => `<brk id="${v}" min="0" max="${m}" man="1"/>`).join("")}</${f}>`
            : "",
        )
        .join("");
      ((d = d.replace(
        /<(?:customProperties|cellWatches|ignoredErrors|smartTags|drawing|legacyDrawing|legacyDrawingHF|picture|oleObjects|controls|webPublishItems|tableParts|extLst)\b|<\/worksheet>/,
        (f) => l + f,
      )),
        (t[s] = strToU8(d)));
    }),
    zipSync(t)
  );
}
function applySheetRuleFlags(e, sheets) {
  if (
    !sheets.some(
      (sheet) =>
        sheet.conditionalFormats?.some(
          (conditionalFormat) => conditionalFormat.stopIfTrue,
        ) ||
        sheet.validations?.some(
          (validation) => validation.showDropdown === false,
        ),
    )
  )
    return e;
  const t = unzipSync(e);
  return (
    sheets.forEach((sheet, o) => {
      const a = `xl/worksheets/sheet${o + 1}.xml`;
      t[a] = strToU8(
        applyValidationDropdownFlags(
          appendStopIfTrueAttribute(strFromU8(t[a]), sheet),
          sheet,
        ),
      );
    }),
    zipSync(t)
  );
}
async function exportWorkbookToXlsx(workbook) {
  const xlsxWorkbookPackage = [
    "Dynamic-array formulas are exported as ordinary anchor formulas without Excel dynamic-array metadata. Native JSON retains local spill behavior. Legacy CSE and metadata-qualified Excel arrays are identified and refused on import; convert a reviewed copy in Excel before importing. Independent Excel round-trip acceptance is required.",
    `Date calculations use the Gregorian ${normalizeDatePolicy(workbook.datePolicy).dateSystem} serial system, preserved in XLSX. Native JSON retains the explicit text-date order and calculation time zone; XLSX does not retain that policy or the calculation clock evidence. DATEVALUE/VALUE text interpretation and TODAY/NOW results may change in Excel. Export calculation evidence separately and review calendar-sensitive formulas.`,
    "XLSX exported with supported values, formulas, styles and page layout. Keep native JSON for editable review history, outline groups, links, pivots, charts, shapes and application metadata. Number formats may differ in Excel. External-action formulas are saved as text.",
  ];
  (workbook.styleCatalog ||
    workbook.sheets.some((sheet) =>
      Object.values(sheet.cells).some((cell) => cell.style?.binding),
    )) &&
    xlsxWorkbookPackage.push(
      "Named styles and theme bindings are flattened to their current effective formatting in XLSX. Keep native JSON to preserve reusable styles, inherited values and explicit overrides.",
    );
  const workbookInput = normalizeWorkbookInput(workbook),
    workbookWithSettings = calculateWorkbookWithSettings(workbookInput);
  (workbookInput.workbookLinks &&
    xlsxWorkbookPackage.push(
      "Workbook link targets, reviewed source caches and last-resolved revision evidence are retained in native JSON only. External formulas export as inert text; use Data > Workbook Links to preview breaking them into values before XLSX export. Retained originals are unchanged.",
    ),
    workbookInput.extensions &&
      xlsxWorkbookPackage.push(
        "Custom-function manifests, reviewed result caches and linked entity fields are retained in native JSON only. XLSX does not install or authorize extensions. Custom function formulas require a compatible, separately approved implementation in the receiving application; cached results are historical and not proof of live calculation.",
      ),
    workbookInput.sourceImports?.length &&
      xlsxWorkbookPackage.push(
        "Structured/document source originals, field paths, native-type evidence, page regions and review decisions are native JSON only and omitted from derived XLSX. Keep a native backup and use File > Import sources to retrieve the originals.",
      ),
    workbookInput.alternateSource &&
      xlsxWorkbookPackage.push(
        "The retained alternate-format original and conversion provenance are native JSON only and are omitted from derived XLSX. Keep the original from File > Workbook formats. Conversion equivalence has not been verified.",
      ),
    workbookInput.reportOutputs?.length &&
      xlsxWorkbookPackage.push(
        "Reusable report outputs, fixed snapshots and host-reviewed links are native JSON only and are omitted from XLSX. Export each report as SVG/PNG/HTML/JSON for downstream reuse; XLSX does not create Office OLE links or refresh authority.",
      ),
    workbookInput.embeddedObjectInventories?.length &&
      xlsxWorkbookPackage.push(
        "Unsupported OLE/ActiveX/embedded-object metadata and payloads are omitted from derived XLSX. Every retained inventory item is listed as dropped in embeddedObjectLosses. No object was executed or fetched; preserve the original files separately and native JSON for the inventory.",
      ),
    workbookInput.sheets.some((sheet) => sheet.analysisReport) &&
      xlsxWorkbookPackage.push(
        "Solver and statistical analysis definitions, methods, assumptions, exclusions and freshness evidence require native JSON. XLSX preserves historical report cells only; native Solver models and Analysis ToolPak settings are not generated. Reopen native JSON to review and rerun.",
      ),
    (workbookInput.dataModel ||
      workbookInput.sheets.some((sheet) => sheet.dataModelReport)) &&
      xlsxWorkbookPackage.push(
        "Relational model tables, relationships, typed contracts, measures, queries and source-freshness evidence require native JSON. XLSX exports supported worksheet values only; no Excel Data Model or DAX is generated. Published model results are snapshots and are not recalculated by Excel from these definitions.",
      ),
    workbookInput.importedCalculationCaches?.entries.length &&
      xlsxWorkbookPackage.push(
        "Original imported formula caches, source file hash and cache comparisons are native JSON evidence only. XLSX writes current local results; it never writes imported historical caches as current values.",
      ),
    workbookInput.calculation?.mode === "manual" &&
      xlsxWorkbookPackage.push(
        "Manual calculation mode is exported. Pending formula caches are omitted; choose Calculate Now before relying on formula results in another application.",
      ),
    (workbookInput.automationArtifacts?.length ||
      workbookInput.automationRuns?.length) &&
      xlsxWorkbookPackage.push(
        "Preserved VBA/Office Scripts source bytes and migration run provenance are native JSON only. XLSX contains the resulting cells; it does not embed executable macros or Office Scripts. Download original artifacts separately from Automation & Scripts.",
      ),
    workbookInput.whatIfAnalyses?.length &&
      xlsxWorkbookPackage.push(
        "What-if scenario and sensitivity definitions, retained assumptions and baseline evidence are native JSON metadata. XLSX preserves report cells only; native Excel Scenario Manager and data-table formulas are not created. Keep native JSON to reopen and refresh the analysis.",
      ),
    workbookInput.sheets.some((sheet) => sheet.pivots?.length) &&
      xlsxWorkbookPackage.push(
        "Native PivotTable definitions, slicers, timelines, Show values as settings, field sorting, PivotCharts, calculated fields, grouping, refresh fingerprints and freshness labels are not preserved in XLSX. Only the current output cells are exported; keep native JSON and export pivot context to verify and refresh the result.",
      ),
    (workbookInput.tableEvidence?.length ||
      workbookInput.sheets.some((sheet) =>
        sheet.tables?.some((table) =>
          table.columns.some((column) => column.dataType),
        ),
      )) &&
      xlsxWorkbookPackage.push(
        "Table column types and original header/conversion evidence are native JSON only. XLSX retains supported tables and current cell values; save native JSON to retain the original source values and type contracts.",
      ),
    workbookInput.textCleanups?.length &&
      xlsxWorkbookPackage.push(
        "Text cleanup history, original evidence snapshots and cleanup recipes are retained in native JSON only. XLSX keeps the output values and cell notes but cannot restore the cleanup history.",
      ),
    workbookInput.savedFilterViews?.length &&
      xlsxWorkbookPackage.push(
        "Named filter views and their historical preview counts are retained in native JSON only. XLSX exports the full source population and does not write Excel AutoFilter criteria. Keep native JSON to reopen saved views and review fresh results.",
      ),
    (workbookInput.transformationRecipes?.length ||
      workbookInput.sheets.some((sheet) => sheet.transformationRun)) &&
      xlsxWorkbookPackage.push(
        "Versioned transformation recipes, source/output hashes and complete step execution evidence are native JSON only. XLSX retains current output values and source-row notes; keep the native workbook for repeatable recipe refresh and lineage.",
      ),
    workbookInput.sheets.some((sheet) => sheet.consolidation) &&
      xlsxWorkbookPackage.push(
        "Consolidation definitions, approved external source identities, review evidence and live source-freshness checks require native JSON. XLSX retains supported cells, same-workbook formulas and source notes, but no consolidation refresh or external workbook links. External amounts remain reviewed snapshots; keep native JSON and reapprove source workbooks to refresh.",
      ),
    workbookInput.sheets.some((sheet) => sheet.externalConnection) &&
      xlsxWorkbookPackage.push(
        "External source connections, refresh attempts, source versions, stale state and control-total evidence are native JSON only. XLSX contains a values snapshot and source notes, with no refresh or credential authority. Keep the native workbook and reauthorize the host source before refreshing.",
      ),
    workbookInput.fieldTransforms?.length &&
      xlsxWorkbookPackage.push(
        "Clean & Reshape recipes, reviewed inference examples and original source snapshots are native JSON only. XLSX keeps current output values, derivation notes and the exceptions appendix; save native JSON to retain reversible transformation evidence.",
      ),
    workbookInput.sheets.some((sheet) =>
      sheet.tables?.some((table) =>
        table.columns.some((column) => column.calculatedFormula),
      ),
    ) &&
      xlsxWorkbookPackage.push(
        "Calculated table column rules are retained in native JSON only. XLSX retains the individual cell formulas and values, but does not preserve the reusable calculated-column rule for future appends.",
      ),
    workbookInput.formulaWatches?.length &&
      xlsxWorkbookPackage.push(
        "The formula Watch Window list is retained in native JSON only; XLSX cellWatches are not exported.",
      ),
    workbookInput.sheets.some((sheet) => sheet.splitPanes) &&
      xlsxWorkbookPackage.push(
        "Split panes are retained in native JSON only and are not preserved in XLSX. Save workbook data to keep divider layouts and initial scroll anchors.",
      ),
    workbookInput.rowOrigins?.length &&
      xlsxWorkbookPackage.push(
        "Original source row identities and their current associations are retained in native JSON only. XLSX export does not preserve the source registry; reopening that XLSX creates a new file-source identity.",
      ),
    workbookInput.definedNames?.length &&
      xlsxWorkbookPackage.push(
        "Defined-name formula text, workbook/worksheet scope and supported comments are preserved. Recalculation uses the supported local engine; relative, volatile and dynamic name formulas require independent Excel review.",
      ),
    workbookInput.sheets.some((sheet) => sheet.delimitedSource) &&
      xlsxWorkbookPackage.push(
        "Delimited source footprints and schemas are not carried by XLSX. Keep native JSON for the historical source footprint and the separate delimited schema report for type choices.",
      ));
  for (const sheet of workbookInput.sheets)
    ((sheet.hiddenRows = cn(sheet, "row")),
      (sheet.hiddenCols = cn(sheet, "column")));
  if (
    workbookInput.sheets.reduce(
      (D, B) =>
        D +
        readHiddenAxisIntervals(B, "row").reduce(
          (_, C) => _ + C.end - C.start + 1,
          0,
        ),
      0,
    ) > worksheetCapacityLimits.maxXlsxHiddenRows
  )
    throw new Error(
      "XLSX export exceeds the 100,000-hidden-row processing budget. Use native JSON to preserve larger hidden ranges without expanding them.",
    );
  const a = new oue.Workbook();
  ((a.creator = "LedgerGlass"),
    (a.properties.date1904 =
      normalizeDatePolicy(workbookInput.datePolicy).dateSystem === "1904"));
  const s = workbookInput.sheets
    .flatMap((sheet) => Object.values(sheet.cells))
    .filter(
      (cell) =>
        typeof cell.raw == "number" &&
        validateXlsxNumericValue(String(cell.raw)),
    );
  (s.length &&
    xlsxWorkbookPackage.push(
      `Numeric precision risk: ${s.length} stored numeric value(s) exceed Excel's 15-digit precision or numeric range. Export preserves stored numbers, not precision as displayed. Use explicit ROUND for monetary amounts and Text for identifiers; retain native JSON.`,
    ),
    workbookInput.metadata &&
      ((a.created = new Date(workbookInput.metadata.createdAt)),
      (a.modified = new Date(workbookInput.metadata.modifiedAt))));
  const exchangeSchemaReport = buildExchangeSchemaReport(
    "export",
    `${workbookInput.name}.xlsx`,
    0,
    workbookInput,
    xlsxWorkbookPackage,
  );
  for (const sheet of workbookInput.sheets) {
    const B = a.addWorksheet(sheet.name, {
      state: sheet.veryHidden
        ? "veryHidden"
        : sheet.hidden
          ? "hidden"
          : "visible",
      properties: {
        defaultRowHeight: defaultRowHeightOf(sheet) * 0.75,
        defaultColWidth: defaultColumnWidthOf(sheet) / 7,
        ...(sheet.tabColor
          ? { tabColor: { argb: `FF${sheet.tabColor.slice(1)}` } }
          : {}),
      },
      views:
        sheet.frozenRows || sheet.frozenCols
          ? [
              {
                state: "frozen",
                ySplit: sheet.frozenRows ?? 0,
                xSplit: sheet.frozenCols ?? 0,
              },
            ]
          : [],
    });
    for (const [_, C] of Object.entries(sheet.columnWidths ?? {}))
      B.getColumn(Number(_) + 1).width = C / 7;
    for (const [_, C] of Object.entries(sheet.rowHeights ?? {}))
      B.getRow(Number(_) + 1).height = C * 0.75;
    for (const _ of readHiddenAxisIntervals(sheet, "column"))
      for (let start = _.start; start <= _.end; start++)
        B.getColumn(start + 1).hidden = true;
    for (const [_, cell] of Object.entries(sheet.cells)) {
      const T = B.getCell(_);
      if (
        typeof cell.raw == "string" &&
        cell.raw.startsWith("=") &&
        !cell.literal &&
        !isExternalActionFormula(cell.raw)
      ) {
        const exportedCellValue = toExportedCellValue(
          workbookWithSettings.getTypedValue(sheet.id, _),
        );
        (exchangeSchemaReport.caches.total++,
          exportedCellValue === undefined
            ? exchangeSchemaReport.caches.omitted++
            : exchangeSchemaReport.caches.written++,
          (T.value = {
            formula: cell.raw.slice(1),
            ...(exportedCellValue !== undefined
              ? { result: exportedCellValue }
              : {}),
          }));
      } else
        ((T.value = cell.raw),
          typeof cell.raw == "string" &&
            cell.raw.startsWith("=") &&
            !cell.literal &&
            (exchangeSchemaReport.caches.total++,
            exchangeSchemaReport.caches.neutralized++));
      if (
        ((T.numFmt = cell.numberFormat
          ? (buildNumberFormatCode(cell.numberFormat) ?? "General")
          : (cell.numFmt ??
            builtinNumberFormatCodes[cell.format ?? "general"])),
        cell.style)
      ) {
        const effectiveCellStyle = resolveEffectiveCellStyle(cell.style);
        ((T.font = {
          bold: effectiveCellStyle.bold,
          italic: effectiveCellStyle.italic,
          strike: effectiveCellStyle.strikethrough,
          underline: effectiveCellStyle.underline,
          name: effectiveCellStyle.fontFamily,
          size: effectiveCellStyle.fontSize,
          ...(effectiveCellStyle.color
            ? {
                color: { argb: expandHexColorToArgb(effectiveCellStyle.color) },
              }
            : {}),
        }),
          effectiveCellStyle.fill &&
            (T.fill = {
              type: "pattern",
              pattern: "solid",
              fgColor: { argb: expandHexColorToArgb(effectiveCellStyle.fill) },
            }),
          (effectiveCellStyle.align ||
            effectiveCellStyle.wrap !== undefined ||
            effectiveCellStyle.verticalAlign ||
            effectiveCellStyle.indent !== undefined) &&
            (T.alignment = {
              horizontal: effectiveCellStyle.align,
              wrapText: effectiveCellStyle.wrap,
              vertical: effectiveCellStyle.verticalAlign,
              indent: effectiveCellStyle.indent,
            }));
        const R = {};
        for (const cellBorderEdge of cellBorderEdges) {
          const borderEdge = resolveBorderEdge(
            effectiveCellStyle,
            cellBorderEdge,
          );
          borderEdge &&
            borderEdge.style !== "none" &&
            (R[cellBorderEdge] = {
              style: borderEdge.style,
              color: {
                argb: expandHexColorToArgb(borderEdge.color ?? "#000000"),
              },
            });
        }
        (effectiveCellStyle.border || effectiveCellStyle.borders) &&
          (T.border = R);
      }
      cell.note && (T.note = cell.note);
    }
    (buildTableCellEntries(B, sheet),
      serializeWorksheetRuleXml(B, sheet),
      xlsxWorkbookPackage.push(...buildDrawingExportNotes(a, B, sheet)),
      xlsxWorkbookPackage.push(
        ...updateSheetPageSetup(B, sheet, workbookInput),
      ));
  }
  a.calcProperties.fullCalcOnLoad = true;
  const u = await a.xlsx.writeBuffer(),
    l = unzipSync(new Uint8Array(u));
  l["xl/workbook.xml"] = strToU8(
    serializeCalculationProperties(
      serializeDefinedNames(strFromU8(l["xl/workbook.xml"]), workbookInput),
      workbookInput,
    ),
  );
  const c = zipSync(l);
  let bytes = new Uint8Array(
    applySheetRuleFlags(
      applyPageBreakXml(
        serializeWorksheetXml(c, workbookInput.sheets),
        workbookInput.sheets,
      ),
      workbookInput.sheets,
    ),
  );
  if (workbookInput.xlsxSource) {
    const D = requireRetainedXlsxSource(workbookInput);
    importXlsxWorkbook(D.bytes, createEmptyExportLedger());
    const B = unzipSync(D.bytes),
      _ = F$(B, D.sha256),
      C = rebuildPreservedPackage(B, unzipSync(bytes), _);
    try {
      if (C.bytes.length > workbookLimits.maxFileBytes)
        throw Error("Preserved metadata exceeds the XLSX reopen budget.");
      (importXlsxWorkbook(C.bytes, createEmptyExportLedger()),
        (bytes = new Uint8Array(C.bytes)),
        (exchangeSchemaReport.partInventory = C.inventory));
    } catch {
      ((exchangeSchemaReport.partInventory = describePreservationDisposition(
        _,
        "Opaque metadata could not pass the final output integrity or capacity checks. Retained only in the original; the derived copy excludes it.",
      )),
        xlsxWorkbookPackage.push(
          "Passive metadata could not be safely included within the export budgets. Inspect the part inventory; download the original for all source content.",
        ));
    }
    xlsxWorkbookPackage.push(
      `${exchangeSchemaReport.partInventory.counts["preserved-opaque"]} source part(s) preserved byte-for-byte as passive metadata. All source bytes remain available through File → Original XLSX and compatibility and native JSON. Derived XLSX does not embed the original source attachment.`,
    );
  }
  const blob = new Blob([bytes], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  if (blob.size > workbookLimits.maxFileBytes)
    throw new Error(
      "XLSX export exceeds the 10 MB reopen limit. Reduce the workbook or use native JSON.",
    );
  return (
    importXlsxWorkbook(bytes, {
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
    }),
    exchangeSchemaReport.caches.omitted &&
      xlsxWorkbookPackage.push(
        `${exchangeSchemaReport.caches.omitted} formula cache(s) omitted because the local result is blank or an application-specific error. Excel must recalculate them; unsupported results are never written as successful text caches.`,
      ),
    exchangeSchemaReport.caches.neutralized &&
      xlsxWorkbookPackage.push(
        `${exchangeSchemaReport.caches.neutralized} external/action formula(s) exported as literal text.`,
      ),
    (exchangeSchemaReport.fileBytes = blob.size),
    (exchangeSchemaReport.warnings = [...new Set(xlsxWorkbookPackage)]),
    {
      blob: blob,
      warnings: exchangeSchemaReport.warnings,
      report: exchangeSchemaReport,
    }
  );
}
function encodeNativeWorkbook(workbook) {
  const r = JSON.stringify(normalizeWorkbookInput(workbook), null, 2);
  if (new TextEncoder().encode(r).length > workbookLimits.maxNativeBytes)
    throw new Error(
      "Native workbook exceeds the 50 MB save limit. Reduce the workbook size before saving.",
    );
  return r;
}

export {
  numericPrecisionPolicy,
  F$,
  retainEmbeddedObjectInventory,
  retainSourceArtifact,
  buildExchangeSchemaReport,
  exportXltxTemplate,
  encodeNativeWorkbook,
  delimitedImportDefaults,
  exportWorksheetToCsv,
  importXlsxWorkbook,
  createEmptyExportLedger,
  exportWorkbookToXlsx,
  verifyAutomationInspection,
  inspectEmbeddedObjects,
  x7,
  getWorksheetUsedRange,
};
export { formatRangeCornerCells, countRangeCells };
