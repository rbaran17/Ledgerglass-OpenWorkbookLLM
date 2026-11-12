import Ox from "saxes";
import { Unzip, UnzipInflate, strToU8, strFromU8, zipSync } from "fflate";
import { iu } from "./01-core-schemas-and-hashing.js";
import { Ja, Mt } from "./02-financial-functions.js";
import { HR, _t, ct } from "./03-workbook-model.js";
import { T$, Yse } from "./05-add-ins-and-custom-functions.js";
import { Jse, Qse, R$, sg } from "./06-analysis-tools.js";

class La extends Error {
  constructor(r, t = sg[r]) {
    (super(t), (this.code = r), (this.name = "FileProcessingError"));
  }
}
const ele = new Set([
    ...Jse,
    "Legacy CSE array formula",
    "Metadata-qualified Excel array formula (dynamic/legacy classification requires Excel metadata)",
  ]),
  pk =
    " Open the original in Excel, repair it if needed, and save a clean, unencrypted XLSX copy.",
  tle = Qse.map((e) => ({
    parts: e,
    expression: new RegExp(
      "^" +
        e
          .map((r) => r.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"))
          .join("([\\s\\S]*?)") +
        "$",
    ),
  }));
function Yw(e, r = 0) {
  if (!(e.length > 4096 || r > 3)) {
    if (ele.has(e)) return e;
    if (e.endsWith(pk)) {
      const t = Yw(e.slice(0, -pk.length), r + 1);
      if (t) return t + pk;
    }
    for (const { parts: t, expression: n } of tle) {
      const o = n.exec(e);
      if (o)
        return t
          .map(
            (a, s) =>
              a +
              (s < t.length - 1
                ? (Yw(o[s + 1], r + 1) ?? "[source detail omitted]")
                : ""),
          )
          .join("");
    }
  }
}
function OC(e) {
  if (e instanceof La) return new La(e.code, Yw(e.message) ?? sg[e.code]);
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
  return new La(t, Yw(r) ?? sg[t]);
}
function rle(e) {
  const r = e,
    t = r && Object.hasOwn(sg, r.code ?? "") ? r.code : "FILE_WORKER",
    n =
      typeof r?.message == "string"
        ? (Yw(r.message) ??
          (Object.values(sg).includes(r.message) ? r.message : sg[t]))
        : sg[t];
  return new La(t, n);
}
const u1 = Object.freeze({
  maxFileBytes: Mt.maxFileBytes,
  maxNativeBytes: Mt.maxNativeBytes,
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
function mk(e, r) {
  return e === "source"
    ? 4 * 1024 * 1024
    : e === "json"
      ? u1.maxConfigurationBytes
      : e === "picture"
        ? u1.maxPictureBytes
        : e === "workbook" && /\.json$/i.test(r)
          ? Mt.maxNativeBytes
          : Mt.maxFileBytes;
}
function c7(e, r, t) {
  var o;
  if (
    typeof r != "string" ||
    !r.length ||
    r.length > 255 ||
    /[\\/\x00-\x1f\x7f\u202a-\u202e\u2066-\u2069]/.test(r)
  )
    throw new La("FILE_TYPE");
  if (e === "workbook" && /\.(xlsm|xlsb|xls|xltm|xlam)$/i.test(r))
    throw new La(
      "FILE_TYPE",
      "Legacy or macro-enabled Excel formats are not supported. Save as XLSX, CSV or native JSON.",
    );
  if (!(
    (o = {
      workbook: /\.(xlsx|xltx|csv|tsv|json)$/i,
      source: /\.(json|xml|pdf|png|jpe?g|html?)$/i,
      formats: /\.(xls|xlsb|xlsm|xltx|xltm|ods)$/i,
      delimited: /\.(csv|tsv)$/i,
      automation: /\.(bas|vba|cls|ts|osts|xlsm|xlam|xltm|xlsx|xltx|xls|xlsb)$/i,
      json: /\.json$/i,
      picture: /\.(png|jpe?g)$/i,
    }[e]) != null && o.test(e === "picture" ? r.trim() : r)
  ))
    throw new La("FILE_TYPE");
  if (!Number.isSafeInteger(t) || t < 0 || t > mk(e, r))
    throw new La(
      "FILE_LIMIT",
      mk(e, r) === Mt.maxNativeBytes
        ? "Native workbook exceeds the 50 MB import limit."
        : mk(e, r) === Mt.maxFileBytes
          ? "Files must be 10 MB or smaller."
          : undefined,
    );
}
function nle(e = {}) {
  const r = e.url === undefined,
    t =
      e.url ?? URL.createObjectURL(new Blob([R$], { type: "text/javascript" }));
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
      (n.onmessage = (a) => {
        return o.onmessage?.call(o, { data: a.data });
      }),
      (n.onerror = () => {
        return o.onerror?.call(o, {});
      }),
      (n.onmessageerror = () => {
        return o.onerror?.call(o, {});
      }),
      o
    );
  } catch (n) {
    throw (r && URL.revokeObjectURL(String(t)), n);
  }
}
async function ile() {
  const e = "node:worker_threads",
    { Worker: r } = await import(e),
    t = new r(
      `const {parentPort}=require("node:worker_threads");globalThis.self=globalThis;globalThis.postMessage=v=>parentPort.postMessage(v);parentPort.on("message",data=>globalThis.onmessage({data}));
` + R$,
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
    t.on("message", (o) => {
      return n.onmessage?.call(n, { data: o });
    }),
    t.on("error", () => {
      return n.onerror?.call(n, {});
    }),
    t.on("messageerror", () => {
      return n.onerror?.call(n, {});
    }),
    t.on("exit", () => {
      return n.onerror?.call(n, {});
    }),
    n
  );
}
function ole() {
  var e;
  if (typeof Worker < "u") return nle();
  if (typeof process < "u" && (e = process.versions) != null && e.node)
    return ile();
  throw Error("File worker unavailable.");
}
const I$ = new WeakMap();
function ale(e, r, t) {
  if (
    e.schemaVersion !== 1 ||
    e.fileName !== t ||
    e.fileBytes !== r.length ||
    e.sha256 !== iu(r)
  )
    throw Error("Invalid isolated inspection.");
  I$.set(e, { signature: JSON.stringify(e), bytes: r });
}
let gk = 0;
function sle(e, r) {
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
            (s(), n(new La("FILE_READ")));
          }),
          (o.onabort = o.onerror),
          r.addEventListener("abort", a, { once: true }),
          o.readAsArrayBuffer(e));
      });
}
async function im(e, r, t = {}, n) {
  var a;
  if (t.reliability) {
    const { reliability: s, ...d } = t;
    return s.run("file", e, () => im(e, r, d, n));
  }
  c7(e, r.name, r.size);
  const o = t.timeoutMs ?? u1.defaultTimeoutMs;
  if (!Number.isInteger(o) || o < 1 || o > u1.maxTimeoutMs)
    throw new La("FILE_LIMIT");
  if ((a = t.signal) != null && a.aborted) throw new La("FILE_CANCELLED");
  if (gk >= u1.maxConcurrent) throw new La("FILE_BUSY");
  return (
    gk++,
    new Promise((s, d) => {
      var g, b;
      let u = false,
        l;
      const c = new AbortController(),
        h = crypto.randomUUID(),
        f = performance.now(),
        p = (y, w) => {
          var E;
          if (!u) {
            if (
              ((u = true),
              gk--,
              clearTimeout(v),
              (E = t.signal) == null || E.removeEventListener("abort", m),
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
        m = () => p(new La("FILE_CANCELLED")),
        v = setTimeout(() => p(new La("FILE_TIMEOUT")), o);
      if (
        ((g = t.signal) == null ||
          g.addEventListener("abort", m, { once: true }),
        (b = t.signal) != null && b.aborted)
      ) {
        m();
        return;
      }
      (async () => {
        let y;
        try {
          y = new Uint8Array(await sle(r, c.signal));
        } catch {
          p(new La("FILE_READ"));
          return;
        }
        if (u) return;
        try {
          c7(e, r.name, y.byteLength);
        } catch (E) {
          p(E);
          return;
        }
        if (y.byteLength !== r.size) {
          p(new La("FILE_READ"));
          return;
        }
        y = y.slice();
        const w =
          e === "automation" || e === "formats" || e === "source"
            ? y.slice()
            : undefined;
        try {
          const E = await (t.workerFactory ?? ole)();
          if (u) {
            try {
              Promise.resolve(E.terminate()).catch(() => {});
            } catch {}
            return;
          }
          ((l = E),
            (l.onerror = () => p(new La("FILE_WORKER"))),
            (l.onmessage = ({ data: A }) => {
              if (u) return;
              const k = A;
              if (!(!k || k.id !== h)) {
                if (performance.now() - f > o) {
                  p(new La("FILE_TIMEOUT"));
                  return;
                }
                if (k.ok !== true) {
                  p(rle(k.error));
                  return;
                }
                if (!Object.hasOwn(k, "value")) {
                  p(new La("FILE_WORKER"));
                  return;
                }
                try {
                  (w && e === "source" && T$(k.value, w, r.name),
                    w && e === "formats" && Yse(k.value, w, r.name),
                    w && e === "automation" && ale(k.value, w, r.name),
                    p(undefined, k.value));
                } catch {
                  p(new La("FILE_WORKER"));
                }
              }
            }),
            l.postMessage(
              { id: h, operation: e, fileName: r.name, bytes: y, options: n },
              [y.buffer],
            ));
        } catch {
          p(new La("FILE_WORKER"));
        }
      })();
    })
  );
}
function H1(e, r) {
  return im("workbook", e, r);
}
function ow(e, r) {
  return im("automation", e, r);
}
function _$(e, r = {}, t) {
  return im("delimited", e, t, r);
}
function MC(e, r) {
  return im("json", e, r);
}
function lle(e, r) {
  return im("picture", e, r);
}
function ule(e, r) {
  return im("formats", e, r);
}

function ple(e, r, t) {
  if (
    t > Ja.maxXlsxExpandedBytes ||
    (r > Ja.xlsxRatioThresholdBytes &&
      r / Math.max(e, 1) > Ja.maxXlsxExpansionRatio)
  )
    throw new Error(
      "The XLSX archive exceeds the safe expansion limit (50 MB or 200:1 ratio).",
    );
}
function mle(e, r, t = 0) {
  const n = (c, h) => {
      return new RegExp(`(?:^|\\s)${h}\\s*=\\s*(["'])(.*?)\\1`).exec(c)?.[2];
    },
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
        ct(m);
      } catch {
        l(
          `Declared dimension ${m} exceeds or violates the supported worksheet bounds.`,
        );
      }
    } else if (f === "row") {
      const m = n(p, "r"),
        v = Number(m);
      ((!m || !/^[1-9]\d*$/.test(m) || v > Ja.maxRows) &&
        l(`Row ${m ?? "(missing)"} must be between 1 and ${Ja.maxRows}.`),
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
        ++u > Ja.maxXlsxCellRecords &&
          l(
            `The XLSX exceeds ${Ja.maxXlsxCellRecords.toLocaleString("en-US")} source cell records, including explicit blank cells.`,
          ));
    }
  }
  return u;
}
const gle =
    "Open the original in Excel, repair it if needed, and save a clean, unencrypted XLSX copy.",
  Zn = (e) => {
    throw Error(`${e} ${gle}`);
  },
  vle = new Uint32Array(256).map((e, r) => {
    for (let t = 0; t < 8; t++) r = r & 1 ? 3988292384 ^ (r >>> 1) : r >>> 1;
    return r >>> 0;
  });
function D$(e, r) {
  for (const t of r) e = vle[(e ^ t) & 255] ^ (e >>> 8);
  return e;
}
function ble(e, r = {}) {
  const t = O$(e);
  return (
    (!t.has("[Content_Types].xml") ||
      (!t.has("xl/workbook.xml") &&
        !(r.allowBinaryWorkbookInventory && t.has("xl/workbook.bin")))) &&
      Zn("This archive is not a supported Excel workbook."),
    t
  );
}
function O$(e) {
  const r = new DataView(e.buffer, e.byteOffset, e.byteLength);
  (e.length >= 8 &&
    r.getUint32(0, true) === 3759263696 &&
    r.getUint32(4, true) === 3776623009 &&
    Zn(
      "Encrypted Office or legacy binary workbooks are unsupported. Remove the open password in the source application.",
    ),
    (e.length < 22 || r.getUint32(0, true) !== 67324752) &&
      Zn("This file is not a supported XLSX archive."));
  let t = -1;
  for (let p = e.length - 22; p >= Math.max(0, e.length - 65557); p--)
    if (
      r.getUint32(p, true) === 101010256 &&
      p + 22 + r.getUint16(p + 20, true) === e.length
    ) {
      t = p;
      break;
    }
  t < 0 && Zn("The XLSX archive is incomplete or has trailing data.");
  const n = r.getUint16(t + 10, true),
    o = r.getUint32(t + 12, true),
    a = r.getUint32(t + 16, true);
  (r.getUint16(t + 4, true) ||
    r.getUint16(t + 6, true) ||
    n !== r.getUint16(t + 8, true) ||
    n === 65535 ||
    n > Ja.maxXlsxEntries ||
    a + o !== t) &&
    Zn("Multi-volume, ZIP64 or oversized XLSX archives are not supported.");
  const s = new Map(),
    d = new Set(),
    u = [],
    l = (p) => {
      try {
        return new TextDecoder("utf-8", { fatal: true }).decode(p);
      } catch {
        return Zn("Unsupported XLSX path encoding.");
      }
    };
  let c = a,
    h = 0;
  for (let p = 0; p < n; p++) {
    (c + 46 > t || r.getUint32(c, true) !== 33639248) &&
      Zn("Invalid XLSX archive directory.");
    const m = r.getUint16(c + 8, true),
      v = r.getUint16(c + 10, true),
      g = r.getUint32(c + 16, true),
      b = r.getUint32(c + 20, true),
      y = r.getUint32(c + 24, true),
      w = r.getUint16(c + 28, true),
      E = r.getUint16(c + 30, true),
      A = r.getUint16(c + 32, true),
      k = r.getUint32(c + 42, true);
    (m & 8257 && Zn("Encrypted XLSX archive entries are unsupported."),
      (m & -2063 ||
        ![0, 8].includes(v) ||
        [b, y, k].includes(4294967295) ||
        r.getUint16(c + 34, true)) &&
        Zn("Unsupported XLSX compression or ZIP64 metadata."),
      (c + 46 + w + E + A > t ||
        k + 30 > a ||
        r.getUint32(k, true) !== 67324752) &&
        Zn("Invalid XLSX archive offsets."));
    const F = l(e.subarray(c + 46, c + 46 + w));
    ((!F ||
      d.has(F.toLowerCase()) ||
      /^[\/]/.test(F) ||
      /[\\\x00-\x1f\x7f%:]/.test(F) ||
      F.split("/").some(
        (D, B, _) => D === ".." || D === "." || (!D && B < _.length - 1),
      )) &&
      Zn("Unsafe or duplicate XLSX archive path."),
      d.add(F.toLowerCase()));
    const I = r.getUint16(k + 26, true),
      O = r.getUint16(k + 28, true),
      j = k + 30 + I + O,
      z = j + b;
    (z > a ||
      r.getUint16(k + 6, true) !== m ||
      r.getUint16(k + 8, true) !== v ||
      l(e.subarray(k + 30, k + 30 + I)) !== F) &&
      Zn("XLSX local headers do not match the archive directory.");
    let V = z;
    if (m & 8) {
      const D = z + (z + 4 <= a && r.getUint32(z, true) === 134695760 ? 4 : 0);
      ((D + 12 > a ||
        r.getUint32(D, true) !== g ||
        r.getUint32(D + 4, true) !== b ||
        r.getUint32(D + 8, true) !== y) &&
        Zn("Invalid XLSX streaming data descriptor."),
        (V = D + 12));
    } else
      (r.getUint32(k + 14, true) !== g ||
        r.getUint32(k + 18, true) !== b ||
        r.getUint32(k + 22, true) !== y) &&
        Zn("XLSX local sizes or checksum do not match the archive directory.");
    (v === 0 && b !== y && Zn("Invalid stored XLSX entry size."),
      u.push({ start: k, end: V }),
      (h += y),
      ple(b, y, h),
      s.set(F, { size: y, crc: g }),
      (c += 46 + w + E + A));
  }
  return (
    u.sort((p, m) => p.start - m.start),
    (c !== t ||
      u.some((p, m) => p.start !== (m ? u[m - 1].end : 0)) ||
      u.at(-1)?.end !== a) &&
      Zn("Overlapping or unlisted XLSX archive entries."),
    s
  );
}
const yle =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
  wle = "http://schemas.openxmlformats.org/package/2006/relationships",
  wk = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/",
  xle = "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
  Ele = () => ({
    warnings: new Set(),
    hasMainType: false,
    hasOfficeDocument: false,
    strings: 0,
    maxString: -1,
    styles: 0,
    maxStyle: -1,
    emptyStringCaches: new Map(),
  });
function r3(e, r) {
  let t;
  try {
    t = decodeURIComponent(r);
  } catch {
    return Zn(`Invalid relationship target in ${e}.`);
  }
  /[\\\x00-\x1f\x7f:?#]/.test(t) && Zn(`Unsafe internal relationship in ${e}.`);
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
      ? (o.length || Zn(`Relationship escapes the XLSX package in ${e}.`),
        o.pop())
      : a && a !== "." && o.push(a);
  return o.join("/");
}
function Cle(e, r, t, n) {
  const o = new Ox.SaxesParser({ xmlns: true });
  let a = 0,
    s = 0;
  const d = new Set(),
    u = [];
  let l;
  const c = /^xl\/worksheets\/[^/]+\.xml$/.test(r);
  (o.on("error", () => Zn(`Malformed XML in ${r}.`)),
    o.on("doctype", () =>
      Zn(`DTD/entity declarations are unsupported in ${r}.`),
    ),
    o.on("cdata", () =>
      Zn(
        `CDATA content is unsupported in ${r}. Resave as standard XLSX text encoding.`,
      ),
    ),
    o.on("opentag", (h) => {
      (++a > 128 || ++s > 1e6) &&
        Zn(`XML complexity exceeds the supported budget in ${r}.`);
      const f = (m) => {
          return h.attributes[m]?.value;
        },
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
          Zn(`Unsupported cell type ${m} in ${r}.`),
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
            Zn(`Invalid cell style index in ${r}.`),
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
            Zn(
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
              (h.local !== "Relationships" || h.uri !== wle)) ||
            (r === "xl/workbook.xml" &&
              (h.local !== "workbook" || h.uri !== xle))) &&
          Zn(
            `Unsupported XLSX package type or Strict XML namespace in ${r}. Save as standard Excel Workbook (.xlsx).`,
          ),
        r === "[Content_Types].xml" && h.local === "Override")
      ) {
        const m = f("PartName"),
          v = f("ContentType");
        ((!m || !v || d.has(m)) &&
          Zn("Invalid or duplicate XLSX content type declaration."),
          d.add(m),
          /model\+data|dataModel/i.test(v) &&
            n.warnings.add(
              "Imported Excel Data Model and DAX content is not executed or translated; model caches are historical and require independent review. Exact bytes remain only in the retained original.",
            ),
          /macroEnabled|vbaProject|activeX|oleObject/i.test(v) &&
            Zn(
              "Macro-enabled or executable workbook content is unsupported; renaming a file does not convert its format.",
            ),
          m === "/xl/workbook.xml" &&
            (v !== yle &&
              Zn(
                "Unsupported workbook content type. Templates and alternate formats must be saved as XLSX.",
              ),
            (n.hasMainType = true)),
          t.has(r3("_rels/.rels", m)) || Zn(`Missing XLSX part ${m}.`));
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
            Zn(`Invalid or duplicate relationship in ${r}.`),
          d.add(m),
          /vbaProject|activeX|oleObject|attachedTemplate/i.test(g) &&
            Zn("Executable or embedded-object relationships are unsupported."),
          b === "External")
        )
          (g !== `${wk}hyperlink` &&
            g !== `${wk}externalLinkPath` &&
            Zn(`Unsupported external relationship in ${r}.`),
            n.warnings.add(
              "External hyperlinks/workbook links were found. No external target is fetched; interactive links and refresh connections are omitted.",
            ));
        else {
          const y = r3(r, v);
          (t.has(y) || Zn(`Missing relationship target ${y} in ${r}.`),
            r === "_rels/.rels" &&
              g === `${wk}officeDocument` &&
              ((n.hasOfficeDocument || y !== "xl/workbook.xml") &&
                Zn("Unsupported or duplicate workbook relationship."),
              (n.hasOfficeDocument = true)));
        }
      }
    }),
    o.on("text", (h) => {
      l != null && l.inValue && (l.value += h);
    }),
    o.on("closetag", (h) => {
      if (l && h.local === "v") {
        const f = l.value;
        if (
          (l.type === "n" &&
            f !== "" &&
            (!/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(f) ||
              !Number.isFinite(Number(f))) &&
            Zn(`Invalid numeric value at ${l.address} in ${r}.`),
          l.type === "b" &&
            !["0", "1"].includes(f) &&
            Zn(
              `Unsupported boolean value at ${l.address} in ${r}; save numeric Boolean encoding in a standard XLSX copy.`,
            ),
          l.type === "s" &&
            ((!/^\d+$/.test(f) || !Number.isSafeInteger(Number(f))) &&
              Zn(`Invalid shared-string index at ${l.address} in ${r}.`),
            (n.maxString = Math.max(n.maxString, Number(f)))),
          l.type === "e" &&
            !/^#(?:NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|SPILL!|CALC!|GETTING_DATA|CONNECT!|BLOCKED!|FIELD!|UNKNOWN!|BUSY!)$/.test(
              f,
            ) &&
            Zn(`Unsupported error value at ${l.address} in ${r}.`),
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
function Ale(e, r) {
  ((!e.hasMainType || !e.hasOfficeDocument) &&
    Zn("Missing standard XLSX workbook content type or relationship."),
    e.maxString >= e.strings &&
      Zn("A worksheet references a missing shared string."),
    e.maxStyle >= e.styles &&
      Zn("A worksheet references a missing cell style."));
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
const kle = Object.freeze(
    HR.map((e) =>
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
  aT = Object.freeze({
    xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
    xltx: "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml",
    xlsm: "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
    xltm: "application/vnd.ms-excel.template.macroEnabled.main+xml",
    xlsb: "application/vnd.ms-excel.sheet.binary.macroEnabled.main",
  });
function Sle(e) {
  if (e.length > Mt.maxFileBytes)
    throw Error("Files must be 10 MB or smaller.");
  const r = O$(e),
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
        if (((l = D$(l, f)), c.push(f), p)) {
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
function Yb(e, r) {
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
function m7(e) {
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
      (Yb(e["META-INF/manifest.xml"], (c) => {
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
      Yb(e["content.xml"], (c) => {
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
  Yb(e["[Content_Types].xml"], (u) => {
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
  const s = Object.entries(aT).find(([, u]) => u === t)?.[0];
  if (
    !s ||
    !e[n] ||
    n !== (s === "xlsb" ? "xl/workbook.bin" : "xl/workbook.xml")
  )
    throw Error("Unsupported alternate workbook content type.");
  if (
    (Yb(e["_rels/.rels"], (u) => {
      const l = r(u);
      if (
        u.uri ===
          "http://schemas.openxmlformats.org/package/2006/relationships" &&
        u.local === "Relationship" &&
        l.Type ===
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
      ) {
        if (l.TargetMode || r3("_rels/.rels", l.Target) !== n)
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
    Yb(e[n], (l) => {
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
function Tle(e, r, t) {
  const n = Sle(e);
  if (m7(n) !== r)
    throw Error(
      "The workbook content does not match the template conversion route.",
    );
  if (
    ((n["[Content_Types].xml"] = strToU8(
      strFromU8(n["[Content_Types].xml"]).replace(aT[r], aT[t]),
    )),
    m7(n) !== t)
  )
    throw Error("The template content type could not be converted safely.");
  return zipSync(n);
}
function Nle(e) {
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
function M$(e, r) {
  const t = e.name.replace(/[^a-zA-Z0-9 _-]/g, "").slice(0, 80) || "Workbook",
    n = [e.xlsxSource?.fileName, e.alternateSource?.fileName].map((d) =>
      d?.toLowerCase(),
    );
  let o = t + "-derived";
  for (; n.includes((o + "." + r).toLowerCase());) o += "-copy";
  return o + "." + r;
}

export {
  Ale,
  Cle,
  D$,
  Ele,
  H1,
  I$,
  La,
  M$,
  MC,
  Nle,
  OC,
  Tle,
  _$,
  ble,
  im,
  kle,
  lle,
  mle,
  ow,
  r3,
  ule,
};
