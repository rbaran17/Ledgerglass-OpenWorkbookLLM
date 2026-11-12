import {
  _t,
  alternateFormatExtensions,
  columnIndexToLetters,
  parseRangeAddress,
  stripFormulaStringLiterals,
  throwWorkbookError,
  validateExtractedSourceTables,
  xlsxInspectionLimits,
} from "./03-workbook-model.js";
import {
  dateSystemMaxSerial,
  dateSystemSerialOffset,
  sha256HexDigest,
} from "./01-schemas-date-functions-and-validation.js";
import {
  excelMessagePatterns,
  fileProcessingErrorMessages,
  unsupportedXlsxMarkers,
} from "./06-unsupported-marker-and-message-catalog.js";
import {
  fileImportLimits,
  workbookLimits,
} from "./02-text-and-financial-functions.js";
import Ox from "saxes";
import { Unzip, UnzipInflate, strFromU8, strToU8, zipSync } from "fflate";

/* From 04-cell-input-paste-and-fill.js */
const formatCellAddress = (e, r) => `${columnIndexToLetters(r)}${e + 1}`;
const maxExcelDateSerial = 2958465;
function utcDateFromParts(e, r, t) {
  const n = new Date(0);
  return (n.setUTCHours(0, 0, 0, 0), n.setUTCFullYear(e, r, t), n);
}

/* From 05-add-ins-and-custom-functions.js */
const sourceInspectionRecords = new WeakMap();
function recordSourceInspection(e, r, fileName) {
  if (
    (validateExtractedSourceTables(e),
    e.fileName !== fileName ||
      e.byteLength !== r.length ||
      e.sha256 !== sha256HexDigest(r))
  )
    throw Error("Source inspection does not match its original.");
  sourceInspectionRecords.set(e, {
    signature: JSON.stringify(e),
    bytes: r.slice(),
  });
}
const formatInspectionRecords = new WeakMap();
function validateIsolatedFormatInspection(e, r, fileName) {
  if (
    !e ||
    e.schemaVersion !== 1 ||
    e.fileName !== fileName ||
    e.fileBytes !== r.length ||
    e.sha256 !== sha256HexDigest(r) ||
    !alternateFormatExtensions.includes(e.format) ||
    !fileName.toLowerCase().endsWith("." + e.format) ||
    e.execution !== "never" ||
    !Array.isArray(e.parts) ||
    e.parts.length > 2e3
  )
    throw Error("Invalid isolated format inspection.");
  formatInspectionRecords.set(e, {
    signature: JSON.stringify(e),
    bytes: r.slice(),
  });
}

/* From 07-file-import-and-worker-dispatch.js */
const unsupportedArrayFormulaReasons = new Set([
  ...unsupportedXlsxMarkers,
  "Legacy CSE array formula",
  "Metadata-qualified Excel array formula (dynamic/legacy classification requires Excel metadata)",
]);
const excelRepairSuffix =
  " Open the original in Excel, repair it if needed, and save a clean, unencrypted XLSX copy.";
const sanitizedMessagePatterns = excelMessagePatterns.map(
  (excelMessagePattern) => ({
    parts: excelMessagePattern,
    expression: new RegExp(
      "^" +
        excelMessagePattern
          .map((r) => r.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"))
          .join("([\\s\\S]*?)") +
        "$",
    ),
  }),
);
function sanitizeFileMessage(e, r = 0) {
  if (!(e.length > 4096 || r > 3)) {
    if (unsupportedArrayFormulaReasons.has(e)) return e;
    if (e.endsWith(excelRepairSuffix)) {
      const t = sanitizeFileMessage(
        e.slice(0, -excelRepairSuffix.length),
        r + 1,
      );
      if (t) return t + excelRepairSuffix;
    }
    for (const { parts, expression } of sanitizedMessagePatterns) {
      const o = expression.exec(e);
      if (o)
        return parts
          .map(
            (part, s) =>
              part +
              (s < parts.length - 1
                ? (sanitizeFileMessage(o[s + 1], r + 1) ??
                  "[source detail omitted]")
                : ""),
          )
          .join("");
    }
  }
}
const u1 = Object.freeze({
  maxFileBytes: workbookLimits.maxFileBytes,
  maxNativeBytes: workbookLimits.maxNativeBytes,
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
function fileByteLimitFor(operation, r) {
  return operation === "source"
    ? 4 * 1024 * 1024
    : operation === "json"
      ? u1.maxConfigurationBytes
      : operation === "picture"
        ? u1.maxPictureBytes
        : operation === "workbook" && /\.json$/i.test(r)
          ? workbookLimits.maxNativeBytes
          : workbookLimits.maxFileBytes;
}
function assertXlsxExpansionLimit(e, size, t) {
  if (
    t > fileImportLimits.maxXlsxExpandedBytes ||
    (size > fileImportLimits.xlsxRatioThresholdBytes &&
      size / Math.max(e, 1) > fileImportLimits.maxXlsxExpansionRatio)
  )
    throw new Error(
      "The XLSX archive exceeds the safe expansion limit (50 MB or 200:1 ratio).",
    );
}
function parseWorksheetCells(e, r, t = 0) {
  const n = (c, h) =>
      new RegExp(`(?:^|\\s)${h}\\s*=\\s*(["'])(.*?)\\1`).exec(c)?.[2],
    o = e.replace(/<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>/g, "");
  let a = 0,
    s = 0,
    d = -1,
    u = t;
  const l = (c) => {
    throw new Error(
      `Worksheet ${r}: ${c} Import stopped; no cells were truncated.`,
    );
  };
  for (const c of o.matchAll(
    /<(\/?)(?:[\w.-]+:)?(dimension|row|c)\b([^>]*)>/g,
  )) {
    const [, h, f, p] = c;
    if (h) {
      f === "row" && (a = 0);
      continue;
    }
    if (f === "dimension") {
      const m = n(p, "ref");
      m || l("The declared dimension has no valid cell range.");
      try {
        parseRangeAddress(m);
      } catch {
        l(
          `Declared dimension ${m} exceeds or violates the supported worksheet bounds.`,
        );
      }
    } else if (f === "row") {
      const m = n(p, "r"),
        v = Number(m);
      ((!m || !/^[1-9]\d*$/.test(m) || v > fileImportLimits.maxRows) &&
        l(
          `Row ${m ?? "(missing)"} must be between 1 and ${fileImportLimits.maxRows}.`,
        ),
        (a || v <= s) && l("Rows must have distinct, ascending row numbers."),
        (s = v),
        (a = /\/\s*$/.test(p) ? 0 : v),
        (d = -1));
    } else {
      const m = n(p, "r");
      (!m || !a) &&
        l("Each cell must have an explicit address inside its source row.");
      let v;
      try {
        v = _t(m);
      } catch {
        l(`Cell ${m} exceeds or violates the supported worksheet bounds.`);
      }
      ((v.row + 1 !== a || v.col <= d) &&
        l(`Cell ${m} has a duplicate, unordered or mismatched row address.`),
        (d = v.col),
        ++u > fileImportLimits.maxXlsxCellRecords &&
          l(
            `The XLSX exceeds ${fileImportLimits.maxXlsxCellRecords.toLocaleString("en-US")} source cell records, including explicit blank cells.`,
          ));
    }
  }
  return u;
}
const excelRepairAdvice =
  "Open the original in Excel, repair it if needed, and save a clean, unencrypted XLSX copy.";
const throwInvalidXlsx = (e) => {
  throw Error(`${e} ${excelRepairAdvice}`);
};
const crc32Table = new Uint32Array(256).map((e, r) => {
  for (let t = 0; t < 8; t++) r = r & 1 ? 3988292384 ^ (r >>> 1) : r >>> 1;
  return r >>> 0;
});
function computeCrc32(e, r) {
  for (const t of r) e = crc32Table[(e ^ t) & 255] ^ (e >>> 8);
  return e;
}
function inspectXlsxPackage(bytes, r = {}) {
  const zipDirectory = parseZipDirectory(bytes);
  return (
    (!zipDirectory.has("[Content_Types].xml") ||
      (!zipDirectory.has("xl/workbook.xml") &&
        !(
          r.allowBinaryWorkbookInventory && zipDirectory.has("xl/workbook.bin")
        ))) &&
      throwInvalidXlsx("This archive is not a supported Excel workbook."),
    zipDirectory
  );
}
function parseZipDirectory(bytes) {
  const r = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  (bytes.length >= 8 &&
    r.getUint32(0, true) === 3759263696 &&
    r.getUint32(4, true) === 3776623009 &&
    throwInvalidXlsx(
      "Encrypted Office or legacy binary workbooks are unsupported. Remove the open password in the source application.",
    ),
    (bytes.length < 22 || r.getUint32(0, true) !== 67324752) &&
      throwInvalidXlsx("This file is not a supported XLSX archive."));
  let t = -1;
  for (let p = bytes.length - 22; p >= Math.max(0, bytes.length - 65557); p--)
    if (
      r.getUint32(p, true) === 101010256 &&
      p + 22 + r.getUint16(p + 20, true) === bytes.length
    ) {
      t = p;
      break;
    }
  t < 0 &&
    throwInvalidXlsx("The XLSX archive is incomplete or has trailing data.");
  const n = r.getUint16(t + 10, true),
    o = r.getUint32(t + 12, true),
    a = r.getUint32(t + 16, true);
  (r.getUint16(t + 4, true) ||
    r.getUint16(t + 6, true) ||
    n !== r.getUint16(t + 8, true) ||
    n === 65535 ||
    n > fileImportLimits.maxXlsxEntries ||
    a + o !== t) &&
    throwInvalidXlsx(
      "Multi-volume, ZIP64 or oversized XLSX archives are not supported.",
    );
  const s = new Map(),
    d = new Set(),
    u = [],
    l = (p) => {
      try {
        return new TextDecoder("utf-8", { fatal: true }).decode(p);
      } catch {
        return throwInvalidXlsx("Unsupported XLSX path encoding.");
      }
    };
  let c = a,
    h = 0;
  for (let p = 0; p < n; p++) {
    (c + 46 > t || r.getUint32(c, true) !== 33639248) &&
      throwInvalidXlsx("Invalid XLSX archive directory.");
    const m = r.getUint16(c + 8, true),
      v = r.getUint16(c + 10, true),
      crc = r.getUint32(c + 16, true),
      b = r.getUint32(c + 20, true),
      size = r.getUint32(c + 24, true),
      w = r.getUint16(c + 28, true),
      E = r.getUint16(c + 30, true),
      A = r.getUint16(c + 32, true),
      start = r.getUint32(c + 42, true);
    (m & 8257 &&
      throwInvalidXlsx("Encrypted XLSX archive entries are unsupported."),
      (m & -2063 ||
        ![0, 8].includes(v) ||
        [b, size, start].includes(4294967295) ||
        r.getUint16(c + 34, true)) &&
        throwInvalidXlsx("Unsupported XLSX compression or ZIP64 metadata."),
      (c + 46 + w + E + A > t ||
        start + 30 > a ||
        r.getUint32(start, true) !== 67324752) &&
        throwInvalidXlsx("Invalid XLSX archive offsets."));
    const F = l(bytes.subarray(c + 46, c + 46 + w));
    ((!F ||
      d.has(F.toLowerCase()) ||
      /^[\/]/.test(F) ||
      /[\\\x00-\x1f\x7f%:]/.test(F) ||
      F.split("/").some(
        (D, B, _) => D === ".." || D === "." || (!D && B < _.length - 1),
      )) &&
      throwInvalidXlsx("Unsafe or duplicate XLSX archive path."),
      d.add(F.toLowerCase()));
    const I = r.getUint16(start + 26, true),
      O = r.getUint16(start + 28, true),
      j = start + 30 + I + O,
      z = j + b;
    (z > a ||
      r.getUint16(start + 6, true) !== m ||
      r.getUint16(start + 8, true) !== v ||
      l(bytes.subarray(start + 30, start + 30 + I)) !== F) &&
      throwInvalidXlsx(
        "XLSX local headers do not match the archive directory.",
      );
    let end = z;
    if (m & 8) {
      const D = z + (z + 4 <= a && r.getUint32(z, true) === 134695760 ? 4 : 0);
      ((D + 12 > a ||
        r.getUint32(D, true) !== crc ||
        r.getUint32(D + 4, true) !== b ||
        r.getUint32(D + 8, true) !== size) &&
        throwInvalidXlsx("Invalid XLSX streaming data descriptor."),
        (end = D + 12));
    } else
      (r.getUint32(start + 14, true) !== crc ||
        r.getUint32(start + 18, true) !== b ||
        r.getUint32(start + 22, true) !== size) &&
        throwInvalidXlsx(
          "XLSX local sizes or checksum do not match the archive directory.",
        );
    (v === 0 &&
      b !== size &&
      throwInvalidXlsx("Invalid stored XLSX entry size."),
      u.push({ start: start, end: end }),
      (h += size),
      assertXlsxExpansionLimit(b, size, h),
      s.set(F, { size: size, crc: crc }),
      (c += 46 + w + E + A));
  }
  return (
    u.sort((p, m) => p.start - m.start),
    (c !== t ||
      u.some((p, m) => p.start !== (m ? u[m - 1].end : 0)) ||
      u.at(-1)?.end !== a) &&
      throwInvalidXlsx("Overlapping or unlisted XLSX archive entries."),
    s
  );
}
const xlsxMainContentType =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
const packageRelationshipsNamespace =
  "http://schemas.openxmlformats.org/package/2006/relationships";
const officeDocumentRelationshipPrefix =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
const spreadsheetMainNamespace =
  "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const createWorkbookParseState = () => ({
  warnings: new Set(),
  hasMainType: false,
  hasOfficeDocument: false,
  strings: 0,
  maxString: -1,
  styles: 0,
  maxStyle: -1,
  emptyStringCaches: new Map(),
});
function resolveRelationshipTarget(e, r) {
  let t;
  try {
    t = decodeURIComponent(r);
  } catch {
    return throwInvalidXlsx(`Invalid relationship target in ${e}.`);
  }
  /[\\\x00-\x1f\x7f:?#]/.test(t) &&
    throwInvalidXlsx(`Unsafe internal relationship in ${e}.`);
  const n =
      e === "_rels/.rels"
        ? []
        : e
            .replace(/_rels\/[^/]+\.rels$/, "")
            .split("/")
            .filter(Boolean),
    o = t.startsWith("/") ? [] : n;
  for (const a of t.split("/"))
    a === ".."
      ? (o.length ||
          throwInvalidXlsx(`Relationship escapes the XLSX package in ${e}.`),
        o.pop())
      : a && a !== "." && o.push(a);
  return o.join("/");
}
function parseWorksheetXml(e, r, t, workbookParseState) {
  const o = new Ox.SaxesParser({ xmlns: true });
  let a = 0,
    s = 0;
  const d = new Set(),
    u = [];
  let l;
  const c = /^xl\/worksheets\/[^/]+\.xml$/.test(r);
  (o.on("error", () => throwInvalidXlsx(`Malformed XML in ${r}.`)),
    o.on("doctype", () =>
      throwInvalidXlsx(`DTD/entity declarations are unsupported in ${r}.`),
    ),
    o.on("cdata", () =>
      throwInvalidXlsx(
        `CDATA content is unsupported in ${r}. Resave as standard XLSX text encoding.`,
      ),
    ),
    o.on("opentag", (h) => {
      (++a > 128 || ++s > 1e6) &&
        throwInvalidXlsx(
          `XML complexity exceeds the supported budget in ${r}.`,
        );
      const f = (m) => h.attributes[m]?.value,
        p = u.at(-1);
      if (
        (u.push(h.local),
        r === "xl/sharedStrings.xml" &&
          h.local === "si" &&
          workbookParseState.strings++,
        r === "xl/styles.xml" &&
          h.local === "xf" &&
          p === "cellXfs" &&
          workbookParseState.styles++,
        c && h.local === "c" && p === "row")
      ) {
        const m = f("t") ?? "n";
        (["n", "s", "str", "inlineStr", "b", "e"].includes(m) ||
          throwInvalidXlsx(`Unsupported cell type ${m} in ${r}.`),
          (l = {
            address: f("r") ?? "unknown cell",
            type: m,
            value: "",
            inValue: false,
            formula: false,
          }));
        const v = f("s");
        v !== undefined &&
          ((!/^\d+$/.test(v) || !Number.isSafeInteger(Number(v))) &&
            throwInvalidXlsx(`Invalid cell style index in ${r}.`),
          (workbookParseState.maxStyle = Math.max(
            workbookParseState.maxStyle,
            Number(v),
          )));
      }
      if (
        (l &&
          p === "c" &&
          h.local === "v" &&
          ((l.value = ""), (l.inValue = true)),
        l &&
          p === "c" &&
          h.local === "f" &&
          ((l.formula = true),
          f("t") &&
            !["normal", "shared", "array"].includes(f("t")) &&
            throwInvalidXlsx(
              `Unsupported formula kind at ${l.address} in ${r}. Convert data-table formulas to reviewed ordinary formulas or values first.`,
            )),
        c &&
          ["sheetProtection", "autoFilter", "extLst", "hyperlinks"].includes(
            h.local,
          ) &&
          workbookParseState.warnings.add(
            `Worksheet ${h.local} metadata was detected. Protection, filter criteria, interactive links and unsupported extensions are not preserved as editable Excel features.`,
          ),
        a === 1 &&
          ((r === "[Content_Types].xml" &&
            (h.local !== "Types" ||
              h.uri !==
                "http://schemas.openxmlformats.org/package/2006/content-types")) ||
            (r.endsWith(".rels") &&
              (h.local !== "Relationships" ||
                h.uri !== packageRelationshipsNamespace)) ||
            (r === "xl/workbook.xml" &&
              (h.local !== "workbook" ||
                h.uri !== spreadsheetMainNamespace))) &&
          throwInvalidXlsx(
            `Unsupported XLSX package type or Strict XML namespace in ${r}. Save as standard Excel Workbook (.xlsx).`,
          ),
        r === "[Content_Types].xml" && h.local === "Override")
      ) {
        const m = f("PartName"),
          v = f("ContentType");
        ((!m || !v || d.has(m)) &&
          throwInvalidXlsx(
            "Invalid or duplicate XLSX content type declaration.",
          ),
          d.add(m),
          /model\+data|dataModel/i.test(v) &&
            workbookParseState.warnings.add(
              "Imported Excel Data Model and DAX content is not executed or translated; model caches are historical and require independent review. Exact bytes remain only in the retained original.",
            ),
          /macroEnabled|vbaProject|activeX|oleObject/i.test(v) &&
            throwInvalidXlsx(
              "Macro-enabled or executable workbook content is unsupported; renaming a file does not convert its format.",
            ),
          m === "/xl/workbook.xml" &&
            (v !== xlsxMainContentType &&
              throwInvalidXlsx(
                "Unsupported workbook content type. Templates and alternate formats must be saved as XLSX.",
              ),
            (workbookParseState.hasMainType = true)),
          t.has(resolveRelationshipTarget("_rels/.rels", m)) ||
            throwInvalidXlsx(`Missing XLSX part ${m}.`));
      }
      if (r.endsWith(".rels") && h.local === "Relationship") {
        const m = f("Id"),
          v = f("Target"),
          g = f("Type"),
          b = f("TargetMode");
        if (
          ((!m ||
            !v ||
            !g ||
            d.has(m) ||
            (b && !["External", "Internal"].includes(b))) &&
            throwInvalidXlsx(`Invalid or duplicate relationship in ${r}.`),
          d.add(m),
          /vbaProject|activeX|oleObject|attachedTemplate/i.test(g) &&
            throwInvalidXlsx(
              "Executable or embedded-object relationships are unsupported.",
            ),
          b === "External")
        )
          (g !== `${officeDocumentRelationshipPrefix}hyperlink` &&
            g !== `${officeDocumentRelationshipPrefix}externalLinkPath` &&
            throwInvalidXlsx(`Unsupported external relationship in ${r}.`),
            workbookParseState.warnings.add(
              "External hyperlinks/workbook links were found. No external target is fetched; interactive links and refresh connections are omitted.",
            ));
        else {
          const relationshipTarget = resolveRelationshipTarget(r, v);
          (t.has(relationshipTarget) ||
            throwInvalidXlsx(
              `Missing relationship target ${relationshipTarget} in ${r}.`,
            ),
            r === "_rels/.rels" &&
              g === `${officeDocumentRelationshipPrefix}officeDocument` &&
              ((workbookParseState.hasOfficeDocument ||
                relationshipTarget !== "xl/workbook.xml") &&
                throwInvalidXlsx(
                  "Unsupported or duplicate workbook relationship.",
                ),
              (workbookParseState.hasOfficeDocument = true)));
        }
      }
    }),
    o.on("text", (h) => {
      l?.inValue && (l.value += h);
    }),
    o.on("closetag", (h) => {
      if (l && h.local === "v") {
        const value = l.value;
        if (
          (l.type === "n" &&
            value !== "" &&
            (!/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(value) ||
              !Number.isFinite(Number(value))) &&
            throwInvalidXlsx(`Invalid numeric value at ${l.address} in ${r}.`),
          l.type === "b" &&
            !["0", "1"].includes(value) &&
            throwInvalidXlsx(
              `Unsupported boolean value at ${l.address} in ${r}; save numeric Boolean encoding in a standard XLSX copy.`,
            ),
          l.type === "s" &&
            ((!/^\d+$/.test(value) || !Number.isSafeInteger(Number(value))) &&
              throwInvalidXlsx(
                `Invalid shared-string index at ${l.address} in ${r}.`,
              ),
            (workbookParseState.maxString = Math.max(
              workbookParseState.maxString,
              Number(value),
            ))),
          l.type === "e" &&
            !/^#(?:NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|SPILL!|CALC!|GETTING_DATA|CONNECT!|BLOCKED!|FIELD!|UNKNOWN!|BUSY!)$/.test(
              value,
            ) &&
            throwInvalidXlsx(
              `Unsupported error value at ${l.address} in ${r}.`,
            ),
          l.type === "e" &&
            !l.formula &&
            workbookParseState.warnings.add(
              "Source constant error cells are flattened to displayed text. Error propagation from those cells is not preserved; review dependent calculations.",
            ),
          l.type === "str" && l.formula && value === "")
        ) {
          const p = workbookParseState.emptyStringCaches.get(r) ?? new Set();
          (p.add(l.address), workbookParseState.emptyStringCaches.set(r, p));
        }
        l.inValue = false;
      }
      (h.local === "c" && (l = undefined), u.pop(), a--);
    }),
    o.write(e).close());
}
function validateWorkbookPackage(workbookParseState, r) {
  ((!workbookParseState.hasMainType || !workbookParseState.hasOfficeDocument) &&
    throwInvalidXlsx(
      "Missing standard XLSX workbook content type or relationship.",
    ),
    workbookParseState.maxString >= workbookParseState.strings &&
      throwInvalidXlsx("A worksheet references a missing shared string."),
    workbookParseState.maxStyle >= workbookParseState.styles &&
      throwInvalidXlsx("A worksheet references a missing cell style."));
  const t = [
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
  for (const n of r)
    for (const [o, a] of t) o.test(n) && workbookParseState.warnings.add(a);
  return [...workbookParseState.warnings];
}
const templateConversionRoutes = Object.freeze(
  alternateFormatExtensions.map((alternateFormatExtension) =>
    Object.freeze({
      format: alternateFormatExtension,
      read:
        alternateFormatExtension === "xltx"
          ? "Native template adapter: supported XLSX subset"
          : "External conversion to XLSX, then reviewed import",
      write:
        alternateFormatExtension === "xltx"
          ? "Derived XLTX template with loss review"
          : "Export reviewed XLSX, then Save As in the source application",
      preserve:
        "Exact original bytes in native JSON; derived output is not lossless",
      formulas:
        alternateFormatExtension === "xltx"
          ? "Supported formulas and cache comparison"
          : "Verify formulas and caches after external conversion",
      dates: "Verify date system, serials and displayed dates",
      styles: "Supported formatting only; compare with original",
      specialParts:
        "Original only unless supported by the XLSX codec; never execute macros or objects",
    }),
  ),
);
const macroEnabledContentTypes = Object.freeze({
  xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
  xltx: "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml",
  xlsm: "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
  xltm: "application/vnd.ms-excel.template.macroEnabled.main+xml",
  xlsb: "application/vnd.ms-excel.sheet.binary.macroEnabled.main",
});
function inspectAlternateWorkbookArchive(bytes) {
  if (bytes.length > workbookLimits.maxFileBytes)
    throw Error("Files must be 10 MB or smaller.");
  const zipDirectory = parseZipDirectory(bytes),
    t = Object.create(null),
    n = new Set();
  let o = 0,
    a;
  const s = new Unzip((d) => {
    if (!zipDirectory.has(d.name) || n.has(d.name)) {
      a = Error("Alternate workbook archive directory mismatch.");
      return;
    }
    n.add(d.name);
    let u = 0,
      l = -1;
    const c = [];
    ((d.ondata = (h, f, p) => {
      if (!a) {
        if (h) {
          a = Error("Corrupt alternate workbook archive.");
          return;
        }
        if (
          ((u += f.length),
          (o += f.length),
          u > zipDirectory.get(d.name).size || o > 50 * 1024 * 1024)
        ) {
          ((a = Error(
            "Alternate workbook archive exceeds its expansion budget.",
          )),
            d.terminate());
          return;
        }
        if (((l = computeCrc32(l, f)), c.push(f), p)) {
          if (
            u !== zipDirectory.get(d.name).size ||
            (l ^ -1) >>> 0 !== zipDirectory.get(d.name).crc
          ) {
            a = Error("Alternate workbook archive checksum mismatch.");
            return;
          }
          const m = new Uint8Array(u);
          let v = 0;
          for (const g of c) (m.set(g, v), (v += g.length));
          t[d.name] = m;
        }
      }
    }),
      d.start());
  });
  s.register(UnzipInflate);
  try {
    s.push(bytes, true);
  } catch {
    throw Error("Unable to inspect corrupt alternate workbook archive.");
  }
  if (a) throw a;
  if (Object.keys(t).length !== zipDirectory.size)
    throw Error("Alternate workbook archive is incomplete.");
  return t;
}
function parseAlternateWorkbookXml(e, r) {
  const t = new TextDecoder("utf-8", { fatal: true }).decode(e);
  if (
    t.includes("\0") ||
    /<\?xml[^>]*encoding\s*=\s*["'](?!utf-8|utf8)[^"']+/i.test(t)
  )
    throw Error("Alternate workbook XML must use UTF-8.");
  const n = new Ox.SaxesParser({ xmlns: true });
  let o = 0,
    a = 0;
  (n.on("doctype", () => {
    throw Error("Alternate workbook XML declarations are unsupported.");
  }),
    n.on("opentag", (s) => {
      if (++o > 128 || ++a > 1e6)
        throw Error("Alternate workbook XML exceeds its inspection budget.");
      r(s);
    }),
    n.on("closetag", () => {
      o--;
    }),
    n.write(t).close());
}
function readOpenDocumentPackage(e) {
  const r = (u) =>
    Object.fromEntries(
      Object.values(u.attributes).map((attribute) => [
        attribute.local,
        attribute.value,
      ]),
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
    let u = false,
      l = false;
    if (
      (parseAlternateWorkbookXml(e["META-INF/manifest.xml"], (c) => {
        const h = r(c);
        if (c.uri === "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0") {
          if (c.local === "encryption-data")
            throw Error(
              "Encrypted OpenDocument workbooks require an unencrypted XLSX conversion.",
            );
          c.local === "file-entry" &&
            h["full-path"] === "/" &&
            h["media-type"] ===
              "application/vnd.oasis.opendocument.spreadsheet" &&
            (u = true);
        }
      }),
      parseAlternateWorkbookXml(e["content.xml"], (c) => {
        c.local === "spreadsheet" &&
          c.uri === "urn:oasis:names:tc:opendocument:xmlns:office:1.0" &&
          (l = true);
      }),
      !u || !l)
    )
      throw Error("Unsupported or mismatched OpenDocument workbook identity.");
    return "ods";
  }
  if (!e["[Content_Types].xml"] || !e["_rels/.rels"])
    throw Error("Unsupported alternate workbook package identity.");
  let t,
    n = "",
    o = 0;
  const a = new Set();
  parseAlternateWorkbookXml(e["[Content_Types].xml"], (u) => {
    if (
      u.uri !== "http://schemas.openxmlformats.org/package/2006/content-types"
    )
      throw Error("Unsupported alternate workbook content-type namespace.");
    const l = r(u);
    if (u.local === "Override") {
      if (!l.PartName || a.has(l.PartName))
        throw Error("Duplicate or missing alternate workbook content type.");
      if (
        (a.add(l.PartName),
        l.PartName === "/xl/workbook.xml" || l.PartName === "/xl/workbook.bin")
      ) {
        if (t) throw Error("Ambiguous alternate workbook identity.");
        ((t = l.ContentType), (n = l.PartName.slice(1)));
      }
    }
  });
  const s = Object.entries(macroEnabledContentTypes).find(
    ([, u]) => u === t,
  )?.[0];
  if (
    !s ||
    !e[n] ||
    n !== (s === "xlsb" ? "xl/workbook.bin" : "xl/workbook.xml")
  )
    throw Error("Unsupported alternate workbook content type.");
  if (
    (parseAlternateWorkbookXml(e["_rels/.rels"], (u) => {
      const l = r(u);
      if (
        u.uri ===
          "http://schemas.openxmlformats.org/package/2006/relationships" &&
        u.local === "Relationship" &&
        l.Type ===
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
      ) {
        if (
          l.TargetMode ||
          resolveRelationshipTarget("_rels/.rels", l.Target) !== n
        )
          throw Error("Invalid alternate workbook main relationship.");
        o++;
      }
    }),
    s === "xlsb")
  ) {
    const u = e[n];
    if (u.length < 3 || u[0] !== 131 || u[1] !== 1 || u[2] !== 0)
      throw Error(
        "Invalid binary workbook beginning record. Convert the original in its source application.",
      );
  } else {
    let u = true;
    parseAlternateWorkbookXml(e[n], (l) => {
      if (
        u &&
        ((u = false),
        l.local !== "workbook" ||
          ![
            "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
            "http://purl.oclc.org/ooxml/spreadsheetml/main",
          ].includes(l.uri))
      )
        throw Error("Invalid alternate workbook XML identity.");
    });
  }
  if (o !== 1) throw Error("Invalid alternate workbook main relationship.");
  return s;
}
function convertTemplateToXlsx(bytes, r, t) {
  const n = inspectAlternateWorkbookArchive(bytes);
  if (readOpenDocumentPackage(n) !== r)
    throw Error(
      "The workbook content does not match the template conversion route.",
    );
  if (
    ((n["[Content_Types].xml"] = strToU8(
      strFromU8(n["[Content_Types].xml"]).replace(
        macroEnabledContentTypes[r],
        macroEnabledContentTypes[t],
      ),
    )),
    readOpenDocumentPackage(n) !== t)
  )
    throw Error("The template content type could not be converted safely.");
  return zipSync(n);
}

/* From 08-xlsx-import-export-and-preservation.js */
const z2 =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
const customPropertiesContentType =
  "application/vnd.openxmlformats-officedocument.custom-properties+xml";
const customXmlPropertiesContentType =
  "application/vnd.openxmlformats-officedocument.customXmlProperties+xml";
const Dle = { scope: "workbook" };
function parseXmlPartStrict(e, r, t, n) {
  const o = new Ox.SaxesParser({ xmlns: true });
  (o.on("opentag", r),
    t && o.on("closetag", t),
    n && o.on("text", n),
    o.on("doctype", () => {
      throw Error("DTD is not permitted in preservation metadata.");
    }),
    o.write(strFromU8(e)).close());
}
function ownerPartForRelsPath(e) {
  return e === "_rels/.rels"
    ? ""
    : e.replace(/(^|\/)\_rels\/([^/]+)\.rels$/, "$1$2");
}
function relationshipsPathFor(e) {
  if (!e) return "_rels/.rels";
  const r = e.lastIndexOf("/");
  return `${e.slice(0, r + 1)}_rels/${e.slice(r + 1)}.rels`;
}
function readPackageRelationships(e) {
  const types = new Map(),
    t = new Map(),
    relationships = [],
    sheets = new Map();
  parseXmlPartStrict(e["[Content_Types].xml"], (a) => {
    (a.local === "Override" &&
      types.set(
        (a.attributes.PartName?.value ?? "").slice(1),
        a.attributes.ContentType?.value ?? "",
      ),
      a.local === "Default" &&
        t.set(
          a.attributes.Extension?.value ?? "",
          a.attributes.ContentType?.value ?? "",
        ));
  });
  for (const [rels, s] of Object.entries(e))
    rels.endsWith(".rels") &&
      parseXmlPartStrict(s, (d) => {
        if (d.local !== "Relationship") return;
        const external = d.attributes.TargetMode?.value === "External",
          l = d.attributes.Target?.value ?? "";
        relationships.push({
          owner: ownerPartForRelsPath(rels),
          rels: rels,
          id: d.attributes.Id?.value ?? "",
          type: d.attributes.Type?.value ?? "",
          external: external,
          target: external ? l : resolveRelationshipTarget(rels, l),
        });
      });
  return (
    parseXmlPartStrict(e["xl/workbook.xml"], (a) => {
      if (a.local === "sheet") {
        const relationship = relationships.find(
          (relationship) =>
            relationship.owner === "xl/workbook.xml" &&
            relationship.id ===
              Object.values(a.attributes).find(
                (attribute) =>
                  attribute.local === "id" && attribute.uri === z2.slice(0, -1),
              )?.value,
        );
        relationship &&
          sheets.set(
            relationship.target,
            a.attributes.name?.value ?? "Unknown worksheet",
          );
      }
    }),
    {
      relationships: relationships,
      sheets: sheets,
      contentType: (a) =>
        types.get(a) ?? t.get(a.split(".").at(-1)) ?? "undeclared",
      types: types,
    }
  );
}
function classifyPreservedPart(path, r) {
  return path.endsWith("/")
    ? "directory"
    : path === "[Content_Types].xml" || path.endsWith(".rels")
      ? "package relationships"
      : /vba|macrosheet|activeX|embedding/i.test(path + r)
        ? "active content"
        : /^customXml\//.test(path)
          ? "custom XML"
          : path === "docProps/custom.xml"
            ? "custom properties"
            : /^xl\/charts\//.test(path)
              ? "charts"
              : /^xl\/(?:model|dataModel)\//i.test(path) ||
                  /model\+data|dataModel/i.test(r)
                ? "Excel Data Model"
                : /^xl\/(pivotTables|pivotCache)\//.test(path)
                  ? "pivots"
                  : /connections|queryTables|externalLinks/.test(path)
                    ? "connections and external links"
                    : /slicers|slicerCaches|timelines/.test(path)
                      ? "slicers and timelines"
                      : /threadedComments|persons/.test(path)
                        ? "threaded comments"
                        : /^_xmlsignatures\//.test(path)
                          ? "signatures"
                          : /^xl\/worksheets\/[^/]+\.xml$/.test(path)
                            ? "worksheet"
                            : path === "xl/workbook.xml"
                              ? "workbook and names"
                              : /^xl\/tables\//.test(path)
                                ? "tables"
                                : /^xl\/drawings\//.test(path)
                                  ? "drawings"
                                  : /^xl\/media\//.test(path)
                                    ? "media"
                                    : /^xl\/comments/.test(path)
                                      ? "notes"
                                      : path === "xl/sharedStrings.xml"
                                        ? "shared strings"
                                        : path === "xl/styles.xml" ||
                                            /^xl\/theme\//.test(path)
                                          ? "formatting"
                                          : /^docProps\//.test(path)
                                            ? "document properties"
                                            : "unknown part";
}
function F$(e, sha256) {
  const packageRelationships = readPackageRelationships(e),
    n = new Map(),
    o = new Map();
  for (const relationship of packageRelationships.relationships)
    (n.set(relationship.owner, [
      ...(n.get(relationship.owner) ?? []),
      relationship,
    ]),
      relationship.external ||
        o.set(relationship.target, [
          ...(o.get(relationship.target) ?? []),
          relationship,
        ]));
  const a = new Map();
  for (const [m, v] of packageRelationships.sheets) {
    const g = [m],
      b = new Set();
    for (; g.length;) {
      const y = g.pop();
      if (b.has(y)) continue;
      b.add(y);
      const w = a.get(y) ?? new Set();
      (w.add(v), a.set(y, w));
      for (const E of n.get(y) ?? []) E.external || g.push(E.target);
    }
  }
  const s = new Set(),
    d = new Map(),
    u = Object.keys(e).filter(
      (m) => m.startsWith("customXml/") && !m.endsWith("/"),
    );
  if (u.length) {
    const m = new Set(u),
      v = packageRelationships.relationships.filter(
        (relationship) =>
          m.has(relationship.owner) ||
          (!relationship.external && m.has(relationship.target)),
      );
    let g = !Object.keys(e).some((E) => /xmlMaps|singleXmlCells/i.test(E));
    ((g &&= u.every(
      (E) =>
        E.endsWith(".rels") ||
        (/\.xml$/.test(E) &&
          [
            "application/xml",
            "text/xml",
            customXmlPropertiesContentType,
          ].includes(packageRelationships.contentType(E))),
    )),
      (g &&= v.every(
        (E) =>
          !E.external &&
          ((E.owner === "xl/workbook.xml" &&
            E.type === z2 + "customXml" &&
            m.has(E.target)) ||
            (m.has(E.owner) &&
              m.has(E.target) &&
              E.type === z2 + "customXmlProps")),
      )),
      (g &&= v.every((E) =>
        E.owner === "xl/workbook.xml"
          ? ["application/xml", "text/xml"].includes(
              packageRelationships.contentType(E.target),
            )
          : packageRelationships.contentType(E.owner) !==
              customXmlPropertiesContentType &&
            packageRelationships.contentType(E.target) ===
              customXmlPropertiesContentType,
      )));
    const b = v.filter((E) => E.owner === "xl/workbook.xml"),
      y = new Set(),
      w = b.map((event) => event.target);
    for (; w.length;) {
      const E = w.pop();
      if (!y.has(E)) {
        (y.add(E),
          e[relationshipsPathFor(E)] && y.add(relationshipsPathFor(E)));
        for (const A of n.get(E) ?? []) A.external || w.push(A.target);
      }
    }
    g &&= b.length > 0 && u.every((E) => y.has(E));
    for (const E of u)
      g
        ? s.add(E)
        : d.set(
            E,
            "Custom XML has bindings, unrecognized types, orphaned parts or relationships outside the passive metadata island. Retained only in the original.",
          );
  }
  const l = "docProps/custom.xml";
  if (e[l]) {
    const m = o.get(l) ?? [];
    let v =
      packageRelationships.contentType(l) === customPropertiesContentType &&
      m.length === 1 &&
      m[0].owner === "" &&
      m[0].type === z2 + "custom-properties" &&
      !n.get(l)?.length;
    (parseXmlPartStrict(e[l], (g) => {
      (g.local === "property" && g.attributes.linkTarget && (v = false),
        (g.uri ===
          "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" &&
          ["Properties", "property"].includes(g.local)) ||
          (g.uri ===
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
            ].includes(g.local)) ||
          (v = false));
    }),
      v
        ? s.add(l)
        : d.set(
            l,
            "Custom properties contain links, unsupported value types or relationships; retained only in the original.",
          ));
  }
  const c = new Set();
  if (e["xl/sharedStrings.xml"]) {
    let m = -1;
    parseXmlPartStrict(e["xl/sharedStrings.xml"], (v) => {
      (v.local === "si" && m++, v.local === "r" && c.add(m));
    });
  }
  let h = 0;
  const parts = Object.entries(e)
    .sort(([m], [v]) => m.localeCompare(v))
    .map(([path, v]) => {
      const family = classifyPreservedPart(
          path,
          packageRelationships.contentType(path),
        ),
        locations = [
          ...(a.get(
            path.endsWith(".rels") ? ownerPartForRelsPath(path) : path,
          ) ?? []),
        ].map((sheet) => ({
          sheet: sheet,
          scope: "sheet",
        }));
      locations.length ||
        locations.push(
          [
            "charts",
            "pivots",
            "drawings",
            "notes",
            "unknown part",
            "media",
          ].includes(family)
            ? {
                scope: "unresolved",
                detail:
                  "No worksheet relationship identifies the affected cells.",
              }
            : Dle,
        );
      let disposition = "dropped",
        reason =
          "No safe editable mapping. Exact source bytes remain in the original file.";
      (family === "Excel Data Model" &&
        (reason =
          "Imported Excel Data Model, relationships and DAX are not executed or translated. Worksheet results are historical snapshots. Exact model bytes remain only in the original source; rebuild reviewed local typed tables and relationships to calculate new results."),
        ["worksheet", "workbook and names", "tables", "notes"].includes(
          family,
        ) &&
          ((disposition = "editable"),
          (reason =
            "Supported content is editable and regenerated; inspect the feature details and conversion notices for losses. XML bytes are not preserved.")),
        [
          "package relationships",
          "shared strings",
          "formatting",
          "document properties",
          "drawings",
        ].includes(family) &&
          ((disposition = "flattened"),
          (reason =
            "Supported content is regenerated or flattened. Unsupported settings and original XML are retained only in the source.")),
        family === "media" &&
          /^image\/(png|jpeg)$/.test(packageRelationships.contentType(path)) &&
          ((disposition = "flattened"),
          (reason =
            "Supported worksheet pictures are re-encoded with their supported placement; orphaned images and chart assets are not carried over.")),
        s.has(path) &&
          ((disposition = "preserved-opaque"),
          (reason =
            "Passive metadata snapshot copied byte-for-byte, without interpreting or refreshing its contents. Its original meaning may be stale after edits.")),
        d.has(path) && (reason = d.get(path)),
        family === "signatures" &&
          (reason =
            "Signatures cannot authenticate an edited workbook; retained only in the original file."));
      const features = [],
        A = (feature, disposition, range, detail, j) => {
          if (++h > xlsxInspectionLimits.maxDetails)
            throw Error(
              "XLSX compatibility details exceed the 100,000-entry budget. Split the source workbook before importing.",
            );
          features.push({
            feature: feature,
            disposition: disposition,
            location: {
              ...locations[0],
              ...(range ? { range: range, scope: "cell" } : {}),
              ...(detail ? { detail: detail } : {}),
            },
            reason:
              j ??
              (disposition === "editable"
                ? "Supported subset is editable; review conversion notices for unsupported attributes."
                : "This source feature is not preserved as an editable feature in the derived workbook."),
          });
        };
      if (
        family === "worksheet" ||
        family === "workbook and names" ||
        family === "tables" ||
        family === "drawings" ||
        family === "pivots" ||
        family === "charts"
      ) {
        let k = "",
          F = "",
          I,
          O = "",
          j = false,
          z = false,
          V = "";
        parseXmlPartStrict(
          v,
          (D) => {
            const B = (_) => D.attributes[_]?.value;
            (D.local === "c" &&
              ((k = B("r") ?? ""),
              (F = B("t") ?? ""),
              F === "e" && A("error cell", "flattened", k)),
              D.local === "v" && ((z = true), (V = "")),
              D.local === "f" && ((j = true), (O = "")),
              D.local === "definedName" &&
                ((I = B("name")), A("defined name", "editable", undefined, I)),
              D.local === "dataValidation" &&
                A("data validation", "editable", B("sqref"), B("type")),
              D.local === "conditionalFormatting" &&
                A("conditional formatting", "editable", B("sqref")),
              D.local === "mergeCell" &&
                A("merged cells", "editable", B("ref")),
              D.local === "table" &&
                A("structured table", "editable", B("ref"), B("name")),
              D.local === "location" && A("pivot output", "dropped", B("ref")),
              D.local === "worksheetSource" &&
                A("pivot source", "dropped", B("ref"), B("sheet")),
              D.local === "r" &&
                k &&
                F === "inlineStr" &&
                A("rich text", "flattened", k),
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
              ].includes(D.local) &&
                A(D.local, "dropped", B("ref"), B("name")));
          },
          (D) => {
            (D.local === "v" &&
              ((z = false),
              k &&
                F === "s" &&
                c.has(Number(V)) &&
                A("rich text", "flattened", k)),
              D.local === "c" && (k = ""),
              D.local === "f" &&
                ((j = false),
                family === "charts"
                  ? A(
                      "chart source reference",
                      "dropped",
                      undefined,
                      O.slice(0, 1024),
                    )
                  : /(?:\[|https?:|file:|WEBSERVICE|HYPERLINK|RTD|DDE|CALL\s*\(|EXEC\s*\()/i.test(
                      O,
                    ) &&
                    A(
                      "external/action formula",
                      "flattened",
                      k,
                      undefined,
                      "External/action references may be neutralized to literal text; the formula cache report provides the actual conversion outcome.",
                    )));
          },
          (D) => {
            (j && (O += D), z && (V += D));
          },
        );
      }
      for (const k of n.get(path) ?? [])
        k.external &&
          A(
            "external relationship",
            "dropped",
            undefined,
            k.type.split("/").at(-1),
            "External target is never fetched or carried into the derived copy; its URI remains only in the original.",
          );
      return {
        path: path,
        contentType: path.endsWith("/")
          ? "directory"
          : packageRelationships.contentType(path),
        bytes: v.length,
        sha256: sha256HexDigest(v),
        family: family,
        disposition: disposition,
        reason: reason,
        locations: locations,
        features: features,
        retainedInOriginal: true,
      };
    });
  return summarizePreservedParts({
    schemaVersion: 1,
    policy: "passive-metadata-v1",
    complete: true,
    phase: "eligible-for-export",
    sourceSha256: sha256,
    parts: parts,
    counts: { editable: 0, "preserved-opaque": 0, flattened: 0, dropped: 0 },
    opaqueBytes: 0,
    note: "All package entries are inventoried. Locations come from source relationships and XML; unresolved locations are explicit. Dispositions describe the derived copy, while all parts remain byte-exact in the retained original. Editable parts can also contain flattened or dropped features. No source part is executed or fetched.",
  });
}
function summarizePreservedParts(e) {
  ((e.counts = {
    editable: 0,
    "preserved-opaque": 0,
    flattened: 0,
    dropped: 0,
  }),
    (e.opaqueBytes = 0));
  for (const part of e.parts)
    (e.counts[part.disposition]++,
      part.disposition === "preserved-opaque" && (e.opaqueBytes += part.bytes));
  return e;
}
const numericPrecisionPolicy = Object.freeze({
  storage: "IEEE-754 binary64",
  excelSignificantDigits: 15,
  rounding:
    "ROUND uses decimal ties away from zero; monetary rounding is explicit, never automatic.",
  display:
    "General displays 15 significant digits. Display precision does not round stored values or calculations.",
});
function validateXlsxNumericValue(e) {
  const r = /^[+-]?(\d+\.?\d*|\.\d+)(?:e([+-]?\d+))?$/i.exec(e.trim());
  if (!r) return;
  const t = r[1].replace(".", "").replace(/^0+/, "").replace(/0+$/, ""),
    n = Number(e);
  if (!Number.isFinite(n)) return "outside the finite numeric range";
  if (t && (n === 0 || Math.abs(n) < 22250738585072014e-324))
    return "below Excel's normal floating-point range";
  if (t.length > 15) return "more than 15 significant digits";
  if (Number.isInteger(n) && !Number.isSafeInteger(n))
    return "outside the exact integer range";
}
function scanXlsxNumericTokens(e) {
  const r = new Map(),
    t = new Ox.SaxesParser({ xmlns: true });
  let n = "",
    o = "",
    a = "",
    s = false;
  t.on("opentag", (u) => {
    u.local === "c"
      ? ((n = String(u.attributes.r?.value ?? "")),
        (o = String(u.attributes.t?.value ?? "n")))
      : u.local === "v" && n && ((s = true), (a = ""));
  });
  const d = (u) => {
    if (s && ((a += u), a.length > 1024))
      throw Error("An XLSX numeric token exceeds 1,024 characters.");
  };
  return (
    t.on("text", d),
    t.on("cdata", d),
    t.on("closetag", (u) => {
      if (u.local === "v") {
        if (o === "n" && a.trim()) {
          const l = a.trim();
          if (
            !/^[+-]?(?:\d+\.?\d*|\.\d+)(?:e[+-]?\d+)?$/i.test(l) ||
            !Number.isFinite(Number(l))
          )
            throw Error(`Invalid finite numeric value at ${n}.`);
          r.set(n, l);
        }
        s = false;
      } else u.local === "c" && (n = "");
    }),
    t.write(e).close(),
    r
  );
}
const formatRangeCornerCells = (e) => {
  const rangeAddress = parseRangeAddress(e);
  return `${formatCellAddress(rangeAddress.firstRow, rangeAddress.firstCol)}:${formatCellAddress(rangeAddress.lastRow, rangeAddress.lastCol)}`;
};
function parseXlsxDataValidations(e) {
  const r = [];
  for (const t of e.matchAll(
    /<(?:\w+:)?dataValidation\b([^>]*?)(?:\/>|>([\s\S]*?)<\/(?:\w+:)?dataValidation>)/g,
  )) {
    const n = (d) =>
        new RegExp("\\b" + d + `=["\\']([^"\\']*)["\\']`).exec(t[1])?.[1],
      o = (n("sqref") ?? "").trim().split(/\s+/).filter(Boolean);
    if (!o.length) continue;
    const anchor = formatRangeCornerCells(o[0]).split(":")[0];
    if (r.length + o.length > 1e3)
      throw Error(
        "XLSX has more than 1,000 validation areas; narrow the source rules before import.",
      );
    const operands = [
      ...(t[2] ?? "").matchAll(
        /<(?:\w+:)?formula[12]\b[^>]*>([\s\S]*?)<\/(?:\w+:)?formula[12]>/g,
      ),
    ].map((d) => d[1].trim());
    for (const d of o)
      r.push({
        range: formatRangeCornerCells(d),
        anchor: anchor,
        operands: operands,
        dropdown: !["1", "true"].includes(n("showDropDown") ?? "0"),
        alert: ["1", "true"].includes(n("showErrorMessage") ?? "0"),
      });
  }
  return r;
}
const dataValidationOperators = {
  between: "between",
  notBetween: "notBetween",
  greater: "greaterThan",
  greaterEqual: "greaterThanOrEqual",
  less: "lessThan",
  lessEqual: "lessThanOrEqual",
  equal: "equal",
  notEqual: "notEqual",
};
const isExternalReferenceFormula = (e) => {
  const r = stripFormulaStringLiterals(e).replace(/"(?:[^"]|"")*"/g, '""');
  return (
    /[\[\]|]/.test(r) ||
    /(?:https?:|file:|javascript:|data:|\\\\)/i.test(r) ||
    /\b(?:WEBSERVICE|HYPERLINK|IMAGE|RTD|DDE|CALL|EXEC|REGISTER(?:\.ID)?|RUN|SQL\.REQUEST|STOCKHISTORY)\s*\(/i.test(
      r.replace(/'(?:[^']|'')*'(?=\s*!)/g, "Sheet"),
    )
  );
};
const quoteFormulaString = (e) => '"' + e.replace(/"/g, '""') + '"';
function buildContainsTextFormula(e, cellAddress) {
  return (
    "ISNUMBER(FIND(LOWER(" +
    quoteFormulaString(e) +
    "),LOWER(" +
    cellAddress +
    '&"")))'
  );
}
function Wle(e, r) {
  const rangeAddress = parseRangeAddress(r),
    cellAddress = formatCellAddress(
      rangeAddress.firstRow,
      rangeAddress.firstCol,
    ),
    o = formatRangeCornerCells(r).replace(/([A-Z]+)(\d+)/g, "$$$1$$$2"),
    a =
      "SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(" +
      cellAddress +
      ',"~","~~"),"*","~*"),"?","~?")';
  return (
    "AND(LEN(" +
    cellAddress +
    ")>0,COUNTIF(" +
    o +
    ',"="&' +
    a +
    ")" +
    (e === "duplicate" ? ">1" : "=1") +
    ")"
  );
}
function Hle(raw) {
  if (raw === null || raw === "") return true;
  if (typeof raw == "number") return Number.isFinite(raw);
  if (typeof raw != "string") return false;
  const r = raw.trim() !== "" && Number.isFinite(Number(raw));
  return !(
    /[^\x20-\x7e]/.test(raw) ||
    /^#(?:NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|ERROR!|CYCLE!|SPILL!|CALC!|LIMIT!|BLOCKED!)/.test(
      raw,
    ) ||
    raw.replace(/[~*?]/g, "~~").length > 254 ||
    (r && String(Number(raw)) !== raw) ||
    (!r && /^[\s+-]*\d/.test(raw))
  );
}
function validateConditionalFormatRules(e) {
  const r = {},
    t = new Set();
  for (const s of e.matchAll(/<(?:\w+:)?cfRule\b([^>]*)>/g)) {
    const d = /\bpriority=["']([^"']*)["']/.exec(s[1])?.[1],
      u = /\bstopIfTrue=["']([^"']+)["']/.exec(s[1])?.[1];
    if (d === undefined && /\bid=["']/.test(s[1])) continue;
    const l = Number(d);
    if (!d || !/^\d+$/.test(d) || !Number.isSafeInteger(l) || l < 1 || t.has(l))
      throw Error("Invalid or duplicate conditional-format priority.");
    if ((t.add(l), d && u !== undefined)) {
      if (!["1", "0", "true", "false"].includes(u))
        throw Error("Invalid conditional-format stop flag.");
      r[String(l)] = u === "1" || u === "true";
    }
  }
  for (const s of e.matchAll(/<(?:\w+:)?cfvo\b([^>]*)>/g)) {
    const d = /\bgte=["']([^"']*)["']/.exec(s[1])?.[1];
    if (d !== undefined && d !== "1" && d !== "true")
      throw Error(
        "Strict or malformed conditional icon thresholds are not supported. Use inclusive thresholds or native JSON.",
      );
  }
  return r;
}
function parseCalculationProperties(e) {
  let r = {
    mode: "auto",
    iterative: false,
    maxIterations: null,
    maxChange: null,
  };
  const t = new Ox.SaxesParser({ xmlns: true });
  let n = false;
  return (
    t.on("opentag", (o) => {
      if (o.local !== "calcPr" || !/spreadsheetml/.test(o.uri)) return;
      if (n)
        throw Error(
          "Duplicate XLSX calculation properties. Repair the source workbook.",
        );
      n = true;
      const a = (h) => {
          const f = o.attributes[h];
          return f && typeof f != "string" ? f.value : f;
        },
        mode = a("calcMode") ?? "auto",
        d = a("iterate"),
        u = a("iterateCount"),
        l = a("iterateDelta"),
        c = a("fullPrecision");
      if (
        !["auto", "autoNoTable", "manual"].includes(mode) ||
        (c !== undefined && !["true", "false", "1", "0"].includes(c)) ||
        (d !== undefined && !["true", "false", "1", "0"].includes(d)) ||
        (u !== undefined &&
          (!/^\d+$/.test(u) ||
            !Number.isSafeInteger(Number(u)) ||
            Number(u) < 1)) ||
        (l !== undefined &&
          (!l.trim() || !Number.isFinite(Number(l)) || Number(l) < 0))
      )
        throw Error(
          "Invalid XLSX calculation properties. Repair the source workbook.",
        );
      r = {
        mode: mode,
        iterative: d === "true" || d === "1",
        maxIterations: u === undefined ? null : Number(u),
        maxChange: l === undefined ? null : Number(l),
        ...(c === undefined
          ? {}
          : { fullPrecision: c === "true" || c === "1" }),
      };
    }),
    t.on("doctype", () => {
      throw Error("XLSX calculation properties cannot contain a DTD.");
    }),
    t.write(e).close(),
    r
  );
}
function flattenEmbeddedObjectItems(workbook) {
  return (workbook.embeddedObjectInventories ?? []).flatMap((r) =>
    r.items.map((item) => ({
      sourceFile: r.fileName,
      sourceSha256: r.sourceSha256,
      path: item.path,
      kind: item.kind,
      bytes: item.bytes,
      ...(item.sha256 ? { sha256: item.sha256 } : {}),
      disposition: "dropped",
      execution: "never",
      reason:
        "Inventory retained in native JSON only; unsupported OLE/ActiveX payload was not imported and is not emitted in derived XLSX.",
    })),
  );
}
function buildExchangeSchemaReport(
  direction,
  fileName,
  fileBytes,
  workbook,
  xlsxWorkbookPackage,
) {
  return {
    schemaVersion: 1,
    direction: direction,
    fileName: fileName,
    fileBytes: fileBytes,
    workbookName: workbook.name,
    revision: workbook.revision,
    compatibility: "derived-subset",
    lossless: false,
    originalFileRetained: !!workbook.xlsxSource,
    ...(workbook.alternateSource
      ? {
          alternateOriginal: {
            fileName: workbook.alternateSource.fileName,
            byteLength: workbook.alternateSource.byteLength,
            sha256: workbook.alternateSource.sha256,
            format: workbook.alternateSource.format,
          },
        }
      : {}),
    ...(workbook.xlsxSource
      ? {
          originalSource: {
            fileName: workbook.xlsxSource.fileName,
            byteLength: workbook.xlsxSource.byteLength,
            sha256: workbook.xlsxSource.sha256,
          },
        }
      : {}),
    sheets: workbook.sheets.map((sheet) => ({
      name: sheet.name,
      storedCells: Object.keys(sheet.cells).length,
      hidden: !!(sheet.hidden || sheet.veryHidden),
    })),
    warnings: [...new Set(xlsxWorkbookPackage)],
    ...(direction === "export" && workbook.embeddedObjectInventories?.length
      ? { embeddedObjectLosses: flattenEmbeddedObjectItems(workbook) }
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
const delimitedImportDefaults = {
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
const delimitedColumnTypes = [
  "auto",
  "text",
  "number",
  "integer",
  "boolean",
  "date",
];
const byteCharacterTable = Array.from({ length: 256 }, (e, r) =>
  String.fromCharCode(r),
);
[
  8364, 129, 8218, 402, 8222, 8230, 8224, 8225, 710, 8240, 352, 8249, 338, 141,
  381, 143, 144, 8216, 8217, 8220, 8221, 8226, 8211, 8212, 732, 8482, 353, 8250,
  339, 157, 382, 376,
].forEach((e, r) => (byteCharacterTable[128 + r] = String.fromCharCode(e)));
const recordDelimitedIssue = (e, record, column, n, o) => {
  (e.issueCount++,
    e.issues.length < 100 &&
      e.issues.push({
        record: record,
        column: column,
        message: n,
        ...(o !== undefined ? { value: o.slice(0, 200) } : {}),
      }));
};
const recordDelimitedChange = (
  e,
  record,
  column,
  address,
  reason,
  before,
  after,
) => {
  (e.changeCount++,
    e.changes.push({
      record: record,
      column: column,
      address: address,
      reason: reason,
      before: before,
      after: after,
    }));
};
function normalizeDelimitedOptions(e = {}) {
  const r = {
    ...delimitedImportDefaults,
    ...e,
    columns: { ...(e.columns ?? {}) },
  };
  for (const [t, n] of Object.entries({
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
    if (!n.includes(String(r[t])))
      throw Error(`Unsupported delimited ${t} setting.`);
  if (typeof r.hasHeader != "boolean" || typeof r.trim != "boolean")
    throw Error("Header and whitespace controls must be boolean.");
  if (r.decimal === r.grouping)
    throw Error("Decimal and grouping separators must differ.");
  for (const [t, n] of Object.entries(r.columns))
    if (
      !/^(0|[1-9]\d*)$/.test(t) ||
      Number(t) >= workbookLimits.maxCols ||
      !delimitedColumnTypes.includes(n)
    )
      throw Error("Invalid column type selection.");
  return r;
}
function buildDelimitedImportReport(direction, fileName, t, delimitedOptions) {
  return {
    schemaVersion: 1,
    direction: direction,
    fileName: fileName,
    fileBytes: t,
    options: delimitedOptions,
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
const escapeRegExpSource = (e) => e.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
function parseDelimitedNumber(raw, delimitedOptions) {
  const t =
      delimitedOptions.grouping === "none" ? "" : delimitedOptions.grouping,
    n = t ? `(?:\\d+|\\d{1,3}(?:${escapeRegExpSource(t)}\\d{3})+)` : "\\d+";
  if (
    !new RegExp(
      `^[+-]?(?:${n}(?:${escapeRegExpSource(delimitedOptions.decimal)}\\d*)?|${escapeRegExpSource(delimitedOptions.decimal)}\\d+)(?:[eE][+-]?\\d+)?$`,
    ).test(raw)
  )
    throw Error("Invalid number for the selected decimal/grouping separators.");
  const a = (t ? raw.split(t).join("") : raw).replace(
      delimitedOptions.decimal,
      ".",
    ),
    s = a
      .split(/[eE]/)[0]
      .replace(/[^0-9]/g, "")
      .replace(/^0+/, "")
      .replace(/0+$/, "").length,
    d = Number(a);
  if (
    !Number.isFinite(d) ||
    s > 15 ||
    (d === 0 && /[1-9]/.test(a.split(/[eE]/)[0]))
  )
    throw Error(
      "Number is outside supported precision/range; select Text to retain every digit.",
    );
  return d;
}
const maxXlsxExpandedByteLimit = fileImportLimits.maxXlsxExpandedBytes;
const xlsxExchangeSupportNotice =
  "XLSX exchange supports cell values, formulas, common formats, supported fonts/strike/alignment/indent/per-edge borders, notes, hidden sheets, column widths, explicit row heights, worksheet default dimensions, bounded hidden rows/columns, frozen rows and columns, basic tables, safe merges, supported validation/conditional-format rules, PNG/JPEG pictures, repeated print titles, full-axis manual print breaks and page setup. XLSX processes at most 100,000 hidden rows per workbook; native JSON keeps larger visibility intervals compactly. Workbook and worksheet-scoped defined names, constant/formula definitions and supported name comments are preserved. Unsupported name attributes are reported; unsafe or malformed definitions stop import/export. Enhanced number formats export supported masks; locale, scaling labels, duration/date-system and accounting alignment fidelity may differ, and enhanced format metadata remains native JSON only. Reusable calculated-column rules are native JSON only; XLSX imports retain supported cell formulas but do not infer rules for future appends. Editable outline groups, threaded discussions and interactive hyperlink metadata remain native-JSON features; XLSX retains only collapsed outline visibility. Native pivot/query refresh definitions, charts/shapes, filters, split panes, default-hidden-row settings, password protection and other unsupported Excel metadata are not preserved. External or executable formulas are converted to literal text on import and export. Keep the original workbook; use native JSON for a lossless application backup.";
const isExternalActionFormula = (e) =>
  /[\[\]|]/.test(
    stripFormulaStringLiterals(e).replace(/"(?:[^"]|"")*"/g, '""'),
  ) ||
  /(?:https?:|file:|javascript:|data:|\\\\)/i.test(e) ||
  /\b(?:_xlfn\.)?(?:WEBSERVICE|HYPERLINK|IMAGE|RTD|DDE|CALL|EXEC|REGISTER(?:\.ID)?|RUN|SQL\.REQUEST|STOCKHISTORY)\s*\(/i.test(
    e,
  );
function getXmlAttribute(e, r) {
  return new RegExp("(?:^|\\s)" + r + `\\s*=\\s*(["'])([\\s\\S]*?)\\1`).exec(
    e,
  )?.[2];
}
function resolvePackagePartPath(e) {
  const r = [];
  for (const t of (e.startsWith("/") ? e.slice(1) : "xl/" + e).split("/"))
    t === ".." ? r.pop() : t !== "." && r.push(t);
  return r.join("/");
}
const decodeXmlEntities = (e) =>
  e.replace(/&(?:amp|lt|gt|quot|apos|#\d+|#x[a-f\d]+);/gi, (r) =>
    r === "&amp;"
      ? "&"
      : r === "&lt;"
        ? "<"
        : r === "&gt;"
          ? ">"
          : r === "&quot;"
            ? '"'
            : r === "&apos;"
              ? "'"
              : String.fromCodePoint(
                  Number.parseInt(
                    r.slice(r[2]?.toLowerCase() === "x" ? 3 : 2, -1),
                    r[2]?.toLowerCase() === "x" ? 16 : 10,
                  ),
                ),
  );
function createEmptyExportLedger() {
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

/* From 04-cell-input-paste-and-fill.js */
const s$ = 864e5;
const excelEpochUtc = Date.UTC(1899, 11, 31);
class DateSeriesRangeError extends Error {}
function throwDateSeriesRangeError() {
  throw new DateSeriesRangeError(
    "Date series must stay between 1899-12-31 and 9999-12-31.",
  );
}
function dateToExcelSerial(e) {
  const r = (e.getTime() - excelEpochUtc) / s$,
    t = r >= 60 ? r + 1 : r;
  return (
    (!Number.isSafeInteger(t) || t < 0 || t > maxExcelDateSerial) &&
      throwDateSeriesRangeError(),
    t
  );
}
function parseSeriesSeedDate(e, r = "1900") {
  if (typeof e == "number")
    return (
      (!Number.isSafeInteger(e) || e < 0 || e > dateSystemMaxSerial(r)) &&
        throwWorkbookError(
          `Date values must be whole Excel ${r} serials between 0 and ${dateSystemMaxSerial(r)}.`,
        ),
      e
    );
  if (
    ((typeof e != "string" || !/^\d{4}-\d{2}-\d{2}$/.test(e)) &&
      throwWorkbookError(
        "Enter a date as YYYY-MM-DD or a whole Excel date serial.",
      ),
    e === "1900-02-29" && r === "1900")
  )
    return 60;
  const [t, n, o] = e.split("-").map(Number),
    a = utcDateFromParts(t, n - 1, o);
  return (
    (a.getUTCFullYear() !== t ||
      a.getUTCMonth() !== n - 1 ||
      a.getUTCDate() !== o) &&
      throwWorkbookError("The series date is not a valid calendar date."),
    parseSeriesSeedDate(dateToExcelSerial(a) - dateSystemSerialOffset(r), r)
  );
}

/* From 07-file-import-and-worker-dispatch.js */
class FileProcessingError extends Error {
  constructor(r, t = fileProcessingErrorMessages[r]) {
    (super(t), (this.code = r), (this.name = "FileProcessingError"));
  }
}
function toFileProcessingError(error) {
  if (error instanceof FileProcessingError)
    return new FileProcessingError(
      error.code,
      sanitizeFileMessage(error.message) ??
        fileProcessingErrorMessages[error.code],
    );
  const r = error instanceof Error ? error.message : "",
    t =
      /exceed|budget|limit|too (?:large|many)|at most|smaller|depth|nesting/i.test(
        r,
      )
        ? "FILE_LIMIT"
        : /macro|active.?content|activex|executable|unsafe|document type|doctype|entity|dangerous|DDE/i.test(
              r,
            )
          ? "FILE_ACTIVE"
          : "FILE_INVALID";
  return new FileProcessingError(
    t,
    sanitizeFileMessage(r) ?? fileProcessingErrorMessages[t],
  );
}
function assertImportFileAllowed(operation, r, t) {
  if (
    typeof r != "string" ||
    !r.length ||
    r.length > 255 ||
    /[\\/\x00-\x1f\x7f\u202a-\u202e\u2066-\u2069]/.test(r)
  )
    throw new FileProcessingError("FILE_TYPE");
  if (operation === "workbook" && /\.(xlsm|xlsb|xls|xltm|xlam)$/i.test(r))
    throw new FileProcessingError(
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
    }[operation]?.test(operation === "picture" ? r.trim() : r)
  )
    throw new FileProcessingError("FILE_TYPE");
  if (!Number.isSafeInteger(t) || t < 0 || t > fileByteLimitFor(operation, r))
    throw new FileProcessingError(
      "FILE_LIMIT",
      fileByteLimitFor(operation, r) === workbookLimits.maxNativeBytes
        ? "Native workbook exceeds the 50 MB import limit."
        : fileByteLimitFor(operation, r) === workbookLimits.maxFileBytes
          ? "Files must be 10 MB or smaller."
          : undefined,
    );
}

/* From 08-xlsx-import-export-and-preservation.js */
function parseDelimitedDateText(raw, dateOrder, t = "1900") {
  const n =
    /^(\d{1,4})([-/])(\d{1,2})\2(\d{1,4})(?:[ T](\d{2}):(\d{2})(?::(\d{2}))?)?$/.exec(
      raw,
    );
  if (!n)
    throw Error(
      "Date must match the selected year/month/day order, optionally HH:mm:ss.",
    );
  const o = [n[1], n[3], n[4]],
    [a, s, d] =
      dateOrder === "ymd"
        ? o
        : dateOrder === "dmy"
          ? [o[2], o[1], o[0]]
          : [o[2], o[0], o[1]];
  if (a.length !== 4) throw Error("Dates require a four-digit year.");
  const u = Number(n[5] ?? 0),
    l = Number(n[6] ?? 0),
    c = Number(n[7] ?? 0);
  if (u > 23 || l > 59 || c > 59) throw Error("Invalid date/time component.");
  return (
    parseSeriesSeedDate(`${a}-${s.padStart(2, "0")}-${d.padStart(2, "0")}`, t) +
    (u * 3600 + l * 60 + c) / 86400
  );
}
function parseDelimitedCellValue(e, r, delimitedOptions, n = "1900") {
  if (e === "") return { raw: delimitedOptions.empty === "text" ? "" : null };
  const raw = delimitedOptions.trim ? e.trim() : e;
  if (r === "auto") {
    if (
      /^-?(?:0|[1-9]\d*)(?:\.\d+)?(?:[eE][+-]?\d+)?$/.test(raw) &&
      raw.split(/[eE]/)[0].replace(/[^0-9]/g, "").length <= 15 &&
      delimitedOptions.decimal === "." &&
      delimitedOptions.grouping === "none"
    )
      try {
        return { raw: parseDelimitedNumber(raw, delimitedOptions) };
      } catch {}
    return { raw: raw, ...(raw.startsWith("=") ? { literal: true } : {}) };
  }
  if (r === "text")
    return { raw: raw, ...(raw.startsWith("=") ? { literal: true } : {}) };
  if (r === "date")
    return {
      raw: parseDelimitedDateText(raw, delimitedOptions.dateOrder, n),
      format: "date",
      numFmt: /[ T]\d\d:/.test(raw) ? "yyyy-mm-dd hh:mm:ss" : "yyyy-mm-dd",
    };
  if (r === "boolean") {
    if (/^(true|false)$/i.test(raw))
      return { raw: raw.toLowerCase() === "true" };
    throw Error("Boolean values must be TRUE or FALSE (case-insensitive).");
  }
  const delimitedNumber = parseDelimitedNumber(raw, delimitedOptions);
  if (r === "integer" && !Number.isSafeInteger(delimitedNumber))
    throw Error("Integer column requires a safe whole number.");
  return { raw: delimitedNumber };
}

export {
  DateSeriesRangeError,
  Dle,
  F$,
  FileProcessingError,
  Hle,
  Wle,
  assertImportFileAllowed,
  assertXlsxExpansionLimit,
  buildContainsTextFormula,
  buildDelimitedImportReport,
  buildExchangeSchemaReport,
  byteCharacterTable,
  classifyPreservedPart,
  computeCrc32,
  convertTemplateToXlsx,
  crc32Table,
  createEmptyExportLedger,
  createWorkbookParseState,
  customPropertiesContentType,
  customXmlPropertiesContentType,
  dataValidationOperators,
  dateToExcelSerial,
  decodeXmlEntities,
  delimitedColumnTypes,
  delimitedImportDefaults,
  escapeRegExpSource,
  excelEpochUtc,
  excelRepairAdvice,
  excelRepairSuffix,
  fileByteLimitFor,
  flattenEmbeddedObjectItems,
  formatCellAddress,
  formatInspectionRecords,
  formatRangeCornerCells,
  getXmlAttribute,
  inspectAlternateWorkbookArchive,
  inspectXlsxPackage,
  isExternalActionFormula,
  isExternalReferenceFormula,
  macroEnabledContentTypes,
  maxExcelDateSerial,
  maxXlsxExpandedByteLimit,
  normalizeDelimitedOptions,
  numericPrecisionPolicy,
  officeDocumentRelationshipPrefix,
  ownerPartForRelsPath,
  packageRelationshipsNamespace,
  parseAlternateWorkbookXml,
  parseCalculationProperties,
  parseDelimitedCellValue,
  parseDelimitedDateText,
  parseDelimitedNumber,
  parseSeriesSeedDate,
  parseWorksheetCells,
  parseWorksheetXml,
  parseXlsxDataValidations,
  parseXmlPartStrict,
  parseZipDirectory,
  quoteFormulaString,
  readOpenDocumentPackage,
  readPackageRelationships,
  recordDelimitedChange,
  recordDelimitedIssue,
  recordSourceInspection,
  relationshipsPathFor,
  resolvePackagePartPath,
  resolveRelationshipTarget,
  s$,
  sanitizeFileMessage,
  sanitizedMessagePatterns,
  scanXlsxNumericTokens,
  sourceInspectionRecords,
  spreadsheetMainNamespace,
  summarizePreservedParts,
  templateConversionRoutes,
  throwDateSeriesRangeError,
  throwInvalidXlsx,
  toFileProcessingError,
  u1,
  unsupportedArrayFormulaReasons,
  utcDateFromParts,
  validateConditionalFormatRules,
  validateIsolatedFormatInspection,
  validateWorkbookPackage,
  validateXlsxNumericValue,
  xlsxExchangeSupportNotice,
  xlsxMainContentType,
  z2,
};
