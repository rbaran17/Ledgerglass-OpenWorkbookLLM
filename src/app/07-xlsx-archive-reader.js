import Ox from "saxes";
import { Unzip, UnzipInflate, strToU8, strFromU8, zipSync } from "fflate";
import { sha256HexDigest } from "./01-core-schemas-and-hashing.js";
import { fileImportLimits, workbookLimits } from "./02-financial-functions.js";
import {
  alternateFormatExtensions,
  _t,
  parseRangeAddress,
} from "./03-workbook-model.js";
import {
  recordSourceInspection,
  validateIsolatedFormatInspection,
} from "./05-add-ins-and-custom-functions.js";
import {
  unsupportedXlsxMarkers,
  excelMessagePatterns,
  fileParserWorkerSource,
  fileProcessingErrorMessages,
} from "./06-analysis-tools.js";

class FileProcessingError extends Error {
  constructor(r, t = fileProcessingErrorMessages[r]) {
    (super(t), (this.code = r), (this.name = "FileProcessingError"));
  }
}
const unsupportedArrayFormulaReasons = new Set([
    ...unsupportedXlsxMarkers,
    "Legacy CSE array formula",
    "Metadata-qualified Excel array formula (dynamic/legacy classification requires Excel metadata)",
  ]),
  excelRepairSuffix =
    " Open the original in Excel, repair it if needed, and save a clean, unencrypted XLSX copy.",
  sanitizedMessagePatterns = excelMessagePatterns.map((e) => ({
    parts: e,
    expression: new RegExp(
      "^" +
        e
          .map((r) => r.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"))
          .join("([\\s\\S]*?)") +
        "$",
    ),
  }));
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
            (a, s) =>
              a +
              (s < parts.length - 1
                ? (sanitizeFileMessage(o[s + 1], r + 1) ??
                  "[source detail omitted]")
                : ""),
          )
          .join("");
    }
  }
}
function toFileProcessingError(e) {
  if (e instanceof FileProcessingError)
    return new FileProcessingError(
      e.code,
      sanitizeFileMessage(e.message) ?? fileProcessingErrorMessages[e.code],
    );
  const r = e instanceof Error ? e.message : "",
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
function workerErrorToFileError(e) {
  const r = e,
    t =
      r && Object.hasOwn(fileProcessingErrorMessages, r.code ?? "")
        ? r.code
        : "FILE_WORKER",
    n =
      typeof r?.message == "string"
        ? (sanitizeFileMessage(r.message) ??
          (Object.values(fileProcessingErrorMessages).includes(r.message)
            ? r.message
            : fileProcessingErrorMessages[t]))
        : fileProcessingErrorMessages[t];
  return new FileProcessingError(t, n);
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
function fileByteLimitFor(e, r) {
  return e === "source"
    ? 4 * 1024 * 1024
    : e === "json"
      ? u1.maxConfigurationBytes
      : e === "picture"
        ? u1.maxPictureBytes
        : e === "workbook" && /\.json$/i.test(r)
          ? workbookLimits.maxNativeBytes
          : workbookLimits.maxFileBytes;
}
function assertImportFileAllowed(e, r, t) {
  if (
    typeof r != "string" ||
    !r.length ||
    r.length > 255 ||
    /[\\/\x00-\x1f\x7f\u202a-\u202e\u2066-\u2069]/.test(r)
  )
    throw new FileProcessingError("FILE_TYPE");
  if (e === "workbook" && /\.(xlsm|xlsb|xls|xltm|xlam)$/i.test(r))
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
    }[e]?.test(e === "picture" ? r.trim() : r)
  )
    throw new FileProcessingError("FILE_TYPE");
  if (!Number.isSafeInteger(t) || t < 0 || t > fileByteLimitFor(e, r))
    throw new FileProcessingError(
      "FILE_LIMIT",
      fileByteLimitFor(e, r) === workbookLimits.maxNativeBytes
        ? "Native workbook exceeds the 50 MB import limit."
        : fileByteLimitFor(e, r) === workbookLimits.maxFileBytes
          ? "Files must be 10 MB or smaller."
          : undefined,
    );
}
function createBrowserFileWorker(e = {}) {
  const r = e.url === undefined,
    t =
      e.url ??
      URL.createObjectURL(
        new Blob([fileParserWorkerSource], { type: "text/javascript" }),
      );
  try {
    const n = new Worker(t, { name: "workbook-file-parser" }),
      o = {
        postMessage: (a, s) => n.postMessage(a, s ?? []),
        terminate: () => {
          (n.terminate(), r && URL.revokeObjectURL(String(t)));
        },
        onmessage: null,
        onerror: null,
      };
    return (
      (n.onmessage = (a) => o.onmessage?.({ data: a.data })),
      (n.onerror = () => o.onerror?.({})),
      (n.onmessageerror = () => o.onerror?.({})),
      o
    );
  } catch (n) {
    throw (r && URL.revokeObjectURL(String(t)), n);
  }
}
async function createNodeFileWorker() {
  const e = "node:worker_threads",
    { Worker } = await import(e),
    t = new Worker(
      `const {parentPort}=require("node:worker_threads");globalThis.self=globalThis;globalThis.postMessage=v=>parentPort.postMessage(v);parentPort.on("message",data=>globalThis.onmessage({data}));
` + fileParserWorkerSource,
      {
        eval: true,
        name: "workbook-file-parser",
        env: {},
        resourceLimits: {
          maxOldGenerationSizeMb: u1.nodeHeapMiB,
          maxYoungGenerationSizeMb: 16,
          stackSizeMb: 4,
        },
      },
    ),
    n = {
      postMessage: (o, a) => t.postMessage(o, a),
      terminate: () => t.terminate(),
      onmessage: null,
      onerror: null,
    };
  return (
    t.on("message", (o) => n.onmessage?.({ data: o })),
    t.on("error", () => n.onerror?.({})),
    t.on("messageerror", () => n.onerror?.({})),
    t.on("exit", () => n.onerror?.({})),
    n
  );
}
function createFileWorker() {
  if (typeof Worker < "u") return createBrowserFileWorker();
  if (typeof process < "u" && process.versions?.node)
    return createNodeFileWorker();
  throw Error("File worker unavailable.");
}
const inspectionSignatureCache = new WeakMap();
function registerIsolatedInspection(e, r, t) {
  if (
    e.schemaVersion !== 1 ||
    e.fileName !== t ||
    e.fileBytes !== r.length ||
    e.sha256 !== sha256HexDigest(r)
  )
    throw Error("Invalid isolated inspection.");
  inspectionSignatureCache.set(e, { signature: JSON.stringify(e), bytes: r });
}
let activeFileImportCount = 0;
function readFileAsArrayBuffer(e, r) {
  return typeof e.arrayBuffer == "function"
    ? e.arrayBuffer()
    : new Promise((t, n) => {
        const o = new FileReader(),
          a = () => o.abort(),
          s = () => r.removeEventListener("abort", a);
        ((o.onload = () => {
          (s(), t(o.result));
        }),
          (o.onerror = () => {
            (s(), n(new FileProcessingError("FILE_READ")));
          }),
          (o.onabort = o.onerror),
          r.addEventListener("abort", a, { once: true }),
          o.readAsArrayBuffer(e));
      });
}
async function runFileImport(e, r, t = {}, n) {
  if (t.reliability) {
    const { reliability, ...d } = t;
    return reliability.run("file", e, () => runFileImport(e, r, d, n));
  }
  assertImportFileAllowed(e, r.name, r.size);
  const o = t.timeoutMs ?? u1.defaultTimeoutMs;
  if (!Number.isInteger(o) || o < 1 || o > u1.maxTimeoutMs)
    throw new FileProcessingError("FILE_LIMIT");
  if (t.signal?.aborted) throw new FileProcessingError("FILE_CANCELLED");
  if (activeFileImportCount >= u1.maxConcurrent)
    throw new FileProcessingError("FILE_BUSY");
  return (
    activeFileImportCount++,
    new Promise((s, d) => {
      let u = false,
        l;
      const c = new AbortController(),
        h = crypto.randomUUID(),
        f = performance.now(),
        p = (y, w) => {
          if (!u) {
            if (
              ((u = true),
              activeFileImportCount--,
              clearTimeout(v),
              t.signal?.removeEventListener("abort", m),
              c.abort(),
              l)
            ) {
              ((l.onmessage = null), (l.onerror = null));
              try {
                Promise.resolve(l.terminate()).catch(() => {});
              } catch {}
            }
            y ? d(y) : s(w);
          }
        },
        m = () => p(new FileProcessingError("FILE_CANCELLED")),
        v = setTimeout(() => p(new FileProcessingError("FILE_TIMEOUT")), o);
      if (
        (t.signal?.addEventListener("abort", m, { once: true }),
        t.signal?.aborted)
      ) {
        m();
        return;
      }
      (async () => {
        let y;
        try {
          y = new Uint8Array(await readFileAsArrayBuffer(r, c.signal));
        } catch {
          p(new FileProcessingError("FILE_READ"));
          return;
        }
        if (u) return;
        try {
          assertImportFileAllowed(e, r.name, y.byteLength);
        } catch (E) {
          p(E);
          return;
        }
        if (y.byteLength !== r.size) {
          p(new FileProcessingError("FILE_READ"));
          return;
        }
        y = y.slice();
        const w =
          e === "automation" || e === "formats" || e === "source"
            ? y.slice()
            : undefined;
        try {
          const E = await (t.workerFactory ?? createFileWorker)();
          if (u) {
            try {
              Promise.resolve(E.terminate()).catch(() => {});
            } catch {}
            return;
          }
          ((l = E),
            (l.onerror = () => p(new FileProcessingError("FILE_WORKER"))),
            (l.onmessage = ({ data }) => {
              if (u) return;
              const k = data;
              if (!(!k || k.id !== h)) {
                if (performance.now() - f > o) {
                  p(new FileProcessingError("FILE_TIMEOUT"));
                  return;
                }
                if (k.ok !== true) {
                  p(workerErrorToFileError(k.error));
                  return;
                }
                if (!Object.hasOwn(k, "value")) {
                  p(new FileProcessingError("FILE_WORKER"));
                  return;
                }
                try {
                  (w &&
                    e === "source" &&
                    recordSourceInspection(k.value, w, r.name),
                    w &&
                      e === "formats" &&
                      validateIsolatedFormatInspection(k.value, w, r.name),
                    w &&
                      e === "automation" &&
                      registerIsolatedInspection(k.value, w, r.name),
                    p(undefined, k.value));
                } catch {
                  p(new FileProcessingError("FILE_WORKER"));
                }
              }
            }),
            l.postMessage(
              { id: h, operation: e, fileName: r.name, bytes: y, options: n },
              [y.buffer],
            ));
        } catch {
          p(new FileProcessingError("FILE_WORKER"));
        }
      })();
    })
  );
}
function importWorkbookFile(e, r) {
  return runFileImport("workbook", e, r);
}
function importAutomationFile(e, r) {
  return runFileImport("automation", e, r);
}
function importDelimitedFile(e, r = {}, t) {
  return runFileImport("delimited", e, t, r);
}
function importNativeJsonFile(e, r) {
  return runFileImport("json", e, r);
}
function importPictureFile(e, r) {
  return runFileImport("picture", e, r);
}
function detectSupportedFileFormats(e, r) {
  return runFileImport("formats", e, r);
}

function assertXlsxExpansionLimit(e, r, t) {
  if (
    t > fileImportLimits.maxXlsxExpandedBytes ||
    (r > fileImportLimits.xlsxRatioThresholdBytes &&
      r / Math.max(e, 1) > fileImportLimits.maxXlsxExpansionRatio)
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
    "Open the original in Excel, repair it if needed, and save a clean, unencrypted XLSX copy.",
  throwInvalidXlsx = (e) => {
    throw Error(`${e} ${excelRepairAdvice}`);
  },
  crc32Table = new Uint32Array(256).map((e, r) => {
    for (let t = 0; t < 8; t++) r = r & 1 ? 3988292384 ^ (r >>> 1) : r >>> 1;
    return r >>> 0;
  });
function computeCrc32(e, r) {
  for (const t of r) e = crc32Table[(e ^ t) & 255] ^ (e >>> 8);
  return e;
}
function inspectXlsxPackage(e, r = {}) {
  const t = parseZipDirectory(e);
  return (
    (!t.has("[Content_Types].xml") ||
      (!t.has("xl/workbook.xml") &&
        !(r.allowBinaryWorkbookInventory && t.has("xl/workbook.bin")))) &&
      throwInvalidXlsx("This archive is not a supported Excel workbook."),
    t
  );
}
function parseZipDirectory(e) {
  const r = new DataView(e.buffer, e.byteOffset, e.byteLength);
  (e.length >= 8 &&
    r.getUint32(0, true) === 3759263696 &&
    r.getUint32(4, true) === 3776623009 &&
    throwInvalidXlsx(
      "Encrypted Office or legacy binary workbooks are unsupported. Remove the open password in the source application.",
    ),
    (e.length < 22 || r.getUint32(0, true) !== 67324752) &&
      throwInvalidXlsx("This file is not a supported XLSX archive."));
  let t = -1;
  for (let p = e.length - 22; p >= Math.max(0, e.length - 65557); p--)
    if (
      r.getUint32(p, true) === 101010256 &&
      p + 22 + r.getUint16(p + 20, true) === e.length
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
      g = r.getUint32(c + 16, true),
      b = r.getUint32(c + 20, true),
      y = r.getUint32(c + 24, true),
      w = r.getUint16(c + 28, true),
      E = r.getUint16(c + 30, true),
      A = r.getUint16(c + 32, true),
      k = r.getUint32(c + 42, true);
    (m & 8257 &&
      throwInvalidXlsx("Encrypted XLSX archive entries are unsupported."),
      (m & -2063 ||
        ![0, 8].includes(v) ||
        [b, y, k].includes(4294967295) ||
        r.getUint16(c + 34, true)) &&
        throwInvalidXlsx("Unsupported XLSX compression or ZIP64 metadata."),
      (c + 46 + w + E + A > t ||
        k + 30 > a ||
        r.getUint32(k, true) !== 67324752) &&
        throwInvalidXlsx("Invalid XLSX archive offsets."));
    const F = l(e.subarray(c + 46, c + 46 + w));
    ((!F ||
      d.has(F.toLowerCase()) ||
      /^[\/]/.test(F) ||
      /[\\\x00-\x1f\x7f%:]/.test(F) ||
      F.split("/").some(
        (D, B, _) => D === ".." || D === "." || (!D && B < _.length - 1),
      )) &&
      throwInvalidXlsx("Unsafe or duplicate XLSX archive path."),
      d.add(F.toLowerCase()));
    const I = r.getUint16(k + 26, true),
      O = r.getUint16(k + 28, true),
      j = k + 30 + I + O,
      z = j + b;
    (z > a ||
      r.getUint16(k + 6, true) !== m ||
      r.getUint16(k + 8, true) !== v ||
      l(e.subarray(k + 30, k + 30 + I)) !== F) &&
      throwInvalidXlsx(
        "XLSX local headers do not match the archive directory.",
      );
    let V = z;
    if (m & 8) {
      const D = z + (z + 4 <= a && r.getUint32(z, true) === 134695760 ? 4 : 0);
      ((D + 12 > a ||
        r.getUint32(D, true) !== g ||
        r.getUint32(D + 4, true) !== b ||
        r.getUint32(D + 8, true) !== y) &&
        throwInvalidXlsx("Invalid XLSX streaming data descriptor."),
        (V = D + 12));
    } else
      (r.getUint32(k + 14, true) !== g ||
        r.getUint32(k + 18, true) !== b ||
        r.getUint32(k + 22, true) !== y) &&
        throwInvalidXlsx(
          "XLSX local sizes or checksum do not match the archive directory.",
        );
    (v === 0 && b !== y && throwInvalidXlsx("Invalid stored XLSX entry size."),
      u.push({ start: k, end: V }),
      (h += y),
      assertXlsxExpansionLimit(b, y, h),
      s.set(F, { size: y, crc: g }),
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
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
  packageRelationshipsNamespace =
    "http://schemas.openxmlformats.org/package/2006/relationships",
  officeDocumentRelationshipPrefix =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/",
  spreadsheetMainNamespace =
    "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
  createWorkbookParseState = () => ({
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
function parseWorksheetXml(e, r, t, n) {
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
        r === "xl/sharedStrings.xml" && h.local === "si" && n.strings++,
        r === "xl/styles.xml" &&
          h.local === "xf" &&
          p === "cellXfs" &&
          n.styles++,
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
          (n.maxStyle = Math.max(n.maxStyle, Number(v))));
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
          n.warnings.add(
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
            n.warnings.add(
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
            (n.hasMainType = true)),
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
            n.warnings.add(
              "External hyperlinks/workbook links were found. No external target is fetched; interactive links and refresh connections are omitted.",
            ));
        else {
          const y = resolveRelationshipTarget(r, v);
          (t.has(y) ||
            throwInvalidXlsx(`Missing relationship target ${y} in ${r}.`),
            r === "_rels/.rels" &&
              g === `${officeDocumentRelationshipPrefix}officeDocument` &&
              ((n.hasOfficeDocument || y !== "xl/workbook.xml") &&
                throwInvalidXlsx(
                  "Unsupported or duplicate workbook relationship.",
                ),
              (n.hasOfficeDocument = true)));
        }
      }
    }),
    o.on("text", (h) => {
      l?.inValue && (l.value += h);
    }),
    o.on("closetag", (h) => {
      if (l && h.local === "v") {
        const f = l.value;
        if (
          (l.type === "n" &&
            f !== "" &&
            (!/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(f) ||
              !Number.isFinite(Number(f))) &&
            throwInvalidXlsx(`Invalid numeric value at ${l.address} in ${r}.`),
          l.type === "b" &&
            !["0", "1"].includes(f) &&
            throwInvalidXlsx(
              `Unsupported boolean value at ${l.address} in ${r}; save numeric Boolean encoding in a standard XLSX copy.`,
            ),
          l.type === "s" &&
            ((!/^\d+$/.test(f) || !Number.isSafeInteger(Number(f))) &&
              throwInvalidXlsx(
                `Invalid shared-string index at ${l.address} in ${r}.`,
              ),
            (n.maxString = Math.max(n.maxString, Number(f)))),
          l.type === "e" &&
            !/^#(?:NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|SPILL!|CALC!|GETTING_DATA|CONNECT!|BLOCKED!|FIELD!|UNKNOWN!|BUSY!)$/.test(
              f,
            ) &&
            throwInvalidXlsx(
              `Unsupported error value at ${l.address} in ${r}.`,
            ),
          l.type === "e" &&
            !l.formula &&
            n.warnings.add(
              "Source constant error cells are flattened to displayed text. Error propagation from those cells is not preserved; review dependent calculations.",
            ),
          l.type === "str" && l.formula && f === "")
        ) {
          const p = n.emptyStringCaches.get(r) ?? new Set();
          (p.add(l.address), n.emptyStringCaches.set(r, p));
        }
        l.inValue = false;
      }
      (h.local === "c" && (l = undefined), u.pop(), a--);
    }),
    o.write(e).close());
}
function validateWorkbookPackage(e, r) {
  ((!e.hasMainType || !e.hasOfficeDocument) &&
    throwInvalidXlsx(
      "Missing standard XLSX workbook content type or relationship.",
    ),
    e.maxString >= e.strings &&
      throwInvalidXlsx("A worksheet references a missing shared string."),
    e.maxStyle >= e.styles &&
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
  for (const n of r) for (const [o, a] of t) o.test(n) && e.warnings.add(a);
  return [...e.warnings];
}
const templateConversionRoutes = Object.freeze(
    alternateFormatExtensions.map((e) =>
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
  macroEnabledContentTypes = Object.freeze({
    xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
    xltx: "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml",
    xlsm: "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
    xltm: "application/vnd.ms-excel.template.macroEnabled.main+xml",
    xlsb: "application/vnd.ms-excel.sheet.binary.macroEnabled.main",
  });
function inspectAlternateWorkbookArchive(e) {
  if (e.length > workbookLimits.maxFileBytes)
    throw Error("Files must be 10 MB or smaller.");
  const r = parseZipDirectory(e),
    t = Object.create(null),
    n = new Set();
  let o = 0,
    a;
  const s = new Unzip((d) => {
    if (!r.has(d.name) || n.has(d.name)) {
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
          u > r.get(d.name).size || o > 50 * 1024 * 1024)
        ) {
          ((a = Error(
            "Alternate workbook archive exceeds its expansion budget.",
          )),
            d.terminate());
          return;
        }
        if (((l = computeCrc32(l, f)), c.push(f), p)) {
          if (
            u !== r.get(d.name).size ||
            (l ^ -1) >>> 0 !== r.get(d.name).crc
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
    s.push(e, true);
  } catch {
    throw Error("Unable to inspect corrupt alternate workbook archive.");
  }
  if (a) throw a;
  if (Object.keys(t).length !== r.size)
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
      Object.values(u.attributes).map((l) => [l.local, l.value]),
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
function convertTemplateToXlsx(e, r, t) {
  const n = inspectAlternateWorkbookArchive(e);
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
function readGeneratedTemplateBuffer(e) {
  return typeof e.arrayBuffer == "function"
    ? e.arrayBuffer()
    : new Promise((r, t) => {
        const n = new FileReader();
        ((n.onload = () => r(n.result)),
          (n.onerror = () =>
            t(Error("Unable to read the generated template copy."))),
          n.readAsArrayBuffer(e));
      });
}
function describeImportSource(e, r) {
  const t = e.name.replace(/[^a-zA-Z0-9 _-]/g, "").slice(0, 80) || "Workbook",
    n = [e.xlsxSource?.fileName, e.alternateSource?.fileName].map((d) =>
      d?.toLowerCase(),
    );
  let o = t + "-derived";
  for (; n.includes((o + "." + r).toLowerCase());) o += "-copy";
  return o + "." + r;
}

export {
  validateWorkbookPackage,
  parseWorksheetXml,
  computeCrc32,
  createWorkbookParseState,
  importWorkbookFile,
  inspectionSignatureCache,
  FileProcessingError,
  describeImportSource,
  importNativeJsonFile,
  readGeneratedTemplateBuffer,
  toFileProcessingError,
  convertTemplateToXlsx,
  importDelimitedFile,
  inspectXlsxPackage,
  runFileImport,
  templateConversionRoutes,
  importPictureFile,
  parseWorksheetCells,
  importAutomationFile,
  resolveRelationshipTarget,
  detectSupportedFileFormats,
};
