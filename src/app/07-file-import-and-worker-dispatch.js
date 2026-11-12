import {
  FileProcessingError,
  assertImportFileAllowed,
  toFileProcessingError,
} from "./06c-package-and-worksheet-parsing.js";

import {
  computeCrc32,
  convertTemplateToXlsx,
  createWorkbookParseState,
  inspectXlsxPackage,
  parseWorksheetCells,
  parseWorksheetXml,
  recordSourceInspection,
  resolveRelationshipTarget,
  sanitizeFileMessage,
  templateConversionRoutes,
  u1,
  validateIsolatedFormatInspection,
  validateWorkbookPackage,
} from "./06c-package-and-worksheet-parsing.js";

import { sha256HexDigest } from "./01-schemas-date-functions-and-validation.js";

import { fileProcessingErrorMessages } from "./06-unsupported-marker-and-message-catalog.js";
import { fileParserWorkerSource } from "./06b-file-parser-worker-source.js";

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
