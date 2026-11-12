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
  } catch (error) {
    throw (r && URL.revokeObjectURL(String(t)), error);
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
function registerIsolatedInspection(e, bytes, t) {
  if (
    e.schemaVersion !== 1 ||
    e.fileName !== t ||
    e.fileBytes !== bytes.length ||
    e.sha256 !== sha256HexDigest(bytes)
  )
    throw Error("Invalid isolated inspection.");
  inspectionSignatureCache.set(e, {
    signature: JSON.stringify(e),
    bytes: bytes,
  });
}
let activeFileImportCount = 0;
function readFileAsArrayBuffer(e, signal) {
  return typeof e.arrayBuffer == "function"
    ? e.arrayBuffer()
    : new Promise((t, n) => {
        const o = new FileReader(),
          a = () => o.abort(),
          s = () => signal.removeEventListener("abort", a);
        ((o.onload = () => {
          (s(), t(o.result));
        }),
          (o.onerror = () => {
            (s(), n(new FileProcessingError("FILE_READ")));
          }),
          (o.onabort = o.onerror),
          signal.addEventListener("abort", a, { once: true }),
          o.readAsArrayBuffer(e));
      });
}
async function runFileImport(operation, r, t = {}, n) {
  if (t.reliability) {
    const { reliability, ...d } = t;
    return reliability.run("file", operation, () =>
      runFileImport(operation, r, d, n),
    );
  }
  assertImportFileAllowed(operation, r.name, r.size);
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
        let bytes;
        try {
          bytes = new Uint8Array(await readFileAsArrayBuffer(r, c.signal));
        } catch {
          p(new FileProcessingError("FILE_READ"));
          return;
        }
        if (u) return;
        try {
          assertImportFileAllowed(operation, r.name, bytes.byteLength);
        } catch (error) {
          p(error);
          return;
        }
        if (bytes.byteLength !== r.size) {
          p(new FileProcessingError("FILE_READ"));
          return;
        }
        bytes = bytes.slice();
        const w =
          operation === "automation" ||
          operation === "formats" ||
          operation === "source"
            ? bytes.slice()
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
                    operation === "source" &&
                    recordSourceInspection(k.value, w, r.name),
                    w &&
                      operation === "formats" &&
                      validateIsolatedFormatInspection(k.value, w, r.name),
                    w &&
                      operation === "automation" &&
                      registerIsolatedInspection(k.value, w, r.name),
                    p(undefined, k.value));
                } catch {
                  p(new FileProcessingError("FILE_WORKER"));
                }
              }
            }),
            l.postMessage(
              {
                id: h,
                operation: operation,
                fileName: r.name,
                bytes: bytes,
                options: n,
              },
              [bytes.buffer],
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
function importDelimitedFile(file, r = {}, t) {
  return runFileImport("delimited", file, t, r);
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

function readGeneratedTemplateBuffer(blob) {
  return typeof blob.arrayBuffer == "function"
    ? blob.arrayBuffer()
    : new Promise((r, t) => {
        const n = new FileReader();
        ((n.onload = () => r(n.result)),
          (n.onerror = () =>
            t(Error("Unable to read the generated template copy."))),
          n.readAsArrayBuffer(blob));
      });
}
function describeImportSource(workbook, r) {
  const t =
      workbook.name.replace(/[^a-zA-Z0-9 _-]/g, "").slice(0, 80) || "Workbook",
    n = [workbook.xlsxSource?.fileName, workbook.alternateSource?.fileName].map(
      (d) => d?.toLowerCase(),
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
