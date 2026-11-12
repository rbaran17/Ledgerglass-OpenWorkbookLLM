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

function workerErrorToFileError(workerError) {
  const rawError = workerError,
    errorCode =
      rawError &&
      Object.hasOwn(fileProcessingErrorMessages, rawError.code ?? "")
        ? rawError.code
        : "FILE_WORKER",
    errorMessage =
      typeof rawError?.message == "string"
        ? (sanitizeFileMessage(rawError.message) ??
          (Object.values(fileProcessingErrorMessages).includes(rawError.message)
            ? rawError.message
            : fileProcessingErrorMessages[errorCode]))
        : fileProcessingErrorMessages[errorCode];
  return new FileProcessingError(errorCode, errorMessage);
}

function createBrowserFileWorker(options = {}) {
  const shouldRevokeUrl = options.url === undefined,
    workerUrl =
      options.url ??
      URL.createObjectURL(
        new Blob([fileParserWorkerSource], { type: "text/javascript" }),
      );
  try {
    const worker = new Worker(workerUrl, { name: "workbook-file-parser" }),
      fileWorker = {
        postMessage: (message, transfer) =>
          worker.postMessage(message, transfer ?? []),
        terminate: () => {
          (worker.terminate(),
            shouldRevokeUrl && URL.revokeObjectURL(String(workerUrl)));
        },
        onmessage: null,
        onerror: null,
      };
    return (
      (worker.onmessage = (messageEvent) =>
        fileWorker.onmessage?.({ data: messageEvent.data })),
      (worker.onerror = () => fileWorker.onerror?.({})),
      (worker.onmessageerror = () => fileWorker.onerror?.({})),
      fileWorker
    );
  } catch (error) {
    throw (shouldRevokeUrl && URL.revokeObjectURL(String(workerUrl)), error);
  }
}
async function createNodeFileWorker() {
  const moduleName = "node:worker_threads",
    { Worker } = await import(moduleName),
    worker = new Worker(
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
    fileWorker = {
      postMessage: (message, transfer) => worker.postMessage(message, transfer),
      terminate: () => worker.terminate(),
      onmessage: null,
      onerror: null,
    };
  return (
    worker.on("message", (messageData) =>
      fileWorker.onmessage?.({ data: messageData }),
    ),
    worker.on("error", () => fileWorker.onerror?.({})),
    worker.on("messageerror", () => fileWorker.onerror?.({})),
    worker.on("exit", () => fileWorker.onerror?.({})),
    fileWorker
  );
}
function createFileWorker() {
  if (typeof Worker < "u") return createBrowserFileWorker();
  if (typeof process < "u" && process.versions?.node)
    return createNodeFileWorker();
  throw Error("File worker unavailable.");
}
const inspectionSignatureCache = new WeakMap();
function registerIsolatedInspection(inspection, bytes, fileName) {
  if (
    inspection.schemaVersion !== 1 ||
    inspection.fileName !== fileName ||
    inspection.fileBytes !== bytes.length ||
    inspection.sha256 !== sha256HexDigest(bytes)
  )
    throw Error("Invalid isolated inspection.");
  inspectionSignatureCache.set(inspection, {
    signature: JSON.stringify(inspection),
    bytes: bytes,
  });
}
let activeFileImportCount = 0;
function readFileAsArrayBuffer(file, signal) {
  return typeof file.arrayBuffer == "function"
    ? file.arrayBuffer()
    : new Promise((resolve, reject) => {
        const reader = new FileReader(),
          abortReader = () => reader.abort(),
          removeAbortListener = () =>
            signal.removeEventListener("abort", abortReader);
        ((reader.onload = () => {
          (removeAbortListener(), resolve(reader.result));
        }),
          (reader.onerror = () => {
            (removeAbortListener(),
              reject(new FileProcessingError("FILE_READ")));
          }),
          (reader.onabort = reader.onerror),
          signal.addEventListener("abort", abortReader, { once: true }),
          reader.readAsArrayBuffer(file));
      });
}
async function runFileImport(
  operation,
  file,
  importOptions = {},
  parserOptions,
) {
  if (importOptions.reliability) {
    const { reliability, ...restOptions } = importOptions;
    return reliability.run("file", operation, () =>
      runFileImport(operation, file, restOptions, parserOptions),
    );
  }
  assertImportFileAllowed(operation, file.name, file.size);
  const timeoutMs = importOptions.timeoutMs ?? u1.defaultTimeoutMs;
  if (
    !Number.isInteger(timeoutMs) ||
    timeoutMs < 1 ||
    timeoutMs > u1.maxTimeoutMs
  )
    throw new FileProcessingError("FILE_LIMIT");
  if (importOptions.signal?.aborted)
    throw new FileProcessingError("FILE_CANCELLED");
  if (activeFileImportCount >= u1.maxConcurrent)
    throw new FileProcessingError("FILE_BUSY");
  return (
    activeFileImportCount++,
    new Promise((resolve, reject) => {
      let isSettled = false,
        fileWorker;
      const abortController = new AbortController(),
        requestId = crypto.randomUUID(),
        startedAt = performance.now(),
        settle = (failure, result) => {
          if (!isSettled) {
            if (
              ((isSettled = true),
              activeFileImportCount--,
              clearTimeout(timeoutId),
              importOptions.signal?.removeEventListener("abort", handleAbort),
              abortController.abort(),
              fileWorker)
            ) {
              ((fileWorker.onmessage = null), (fileWorker.onerror = null));
              try {
                Promise.resolve(fileWorker.terminate()).catch(() => {});
              } catch {}
            }
            failure ? reject(failure) : resolve(result);
          }
        },
        handleAbort = () => settle(new FileProcessingError("FILE_CANCELLED")),
        timeoutId = setTimeout(
          () => settle(new FileProcessingError("FILE_TIMEOUT")),
          timeoutMs,
        );
      if (
        (importOptions.signal?.addEventListener("abort", handleAbort, {
          once: true,
        }),
        importOptions.signal?.aborted)
      ) {
        handleAbort();
        return;
      }
      (async () => {
        let bytes;
        try {
          bytes = new Uint8Array(
            await readFileAsArrayBuffer(file, abortController.signal),
          );
        } catch {
          settle(new FileProcessingError("FILE_READ"));
          return;
        }
        if (isSettled) return;
        try {
          assertImportFileAllowed(operation, file.name, bytes.byteLength);
        } catch (error) {
          settle(error);
          return;
        }
        if (bytes.byteLength !== file.size) {
          settle(new FileProcessingError("FILE_READ"));
          return;
        }
        bytes = bytes.slice();
        const isolatedBytes =
          operation === "automation" ||
          operation === "formats" ||
          operation === "source"
            ? bytes.slice()
            : undefined;
        try {
          const worker = await (
            importOptions.workerFactory ?? createFileWorker
          )();
          if (isSettled) {
            try {
              Promise.resolve(worker.terminate()).catch(() => {});
            } catch {}
            return;
          }
          ((fileWorker = worker),
            (fileWorker.onerror = () =>
              settle(new FileProcessingError("FILE_WORKER"))),
            (fileWorker.onmessage = ({ data }) => {
              if (isSettled) return;
              const response = data;
              if (!(!response || response.id !== requestId)) {
                if (performance.now() - startedAt > timeoutMs) {
                  settle(new FileProcessingError("FILE_TIMEOUT"));
                  return;
                }
                if (response.ok !== true) {
                  settle(workerErrorToFileError(response.error));
                  return;
                }
                if (!Object.hasOwn(response, "value")) {
                  settle(new FileProcessingError("FILE_WORKER"));
                  return;
                }
                try {
                  (isolatedBytes &&
                    operation === "source" &&
                    recordSourceInspection(
                      response.value,
                      isolatedBytes,
                      file.name,
                    ),
                    isolatedBytes &&
                      operation === "formats" &&
                      validateIsolatedFormatInspection(
                        response.value,
                        isolatedBytes,
                        file.name,
                      ),
                    isolatedBytes &&
                      operation === "automation" &&
                      registerIsolatedInspection(
                        response.value,
                        isolatedBytes,
                        file.name,
                      ),
                    settle(undefined, response.value));
                } catch {
                  settle(new FileProcessingError("FILE_WORKER"));
                }
              }
            }),
            fileWorker.postMessage(
              {
                id: requestId,
                operation: operation,
                fileName: file.name,
                bytes: bytes,
                options: parserOptions,
              },
              [bytes.buffer],
            ));
        } catch {
          settle(new FileProcessingError("FILE_WORKER"));
        }
      })();
    })
  );
}
function importWorkbookFile(file, importOptions) {
  return runFileImport("workbook", file, importOptions);
}
function importAutomationFile(file, importOptions) {
  return runFileImport("automation", file, importOptions);
}
function importDelimitedFile(file, parserOptions = {}, importOptions) {
  return runFileImport("delimited", file, importOptions, parserOptions);
}
function importNativeJsonFile(file, importOptions) {
  return runFileImport("json", file, importOptions);
}
function importPictureFile(file, importOptions) {
  return runFileImport("picture", file, importOptions);
}
function detectSupportedFileFormats(file, importOptions) {
  return runFileImport("formats", file, importOptions);
}

function readGeneratedTemplateBuffer(blob) {
  return typeof blob.arrayBuffer == "function"
    ? blob.arrayBuffer()
    : new Promise((resolve, reject) => {
        const reader = new FileReader();
        ((reader.onload = () => resolve(reader.result)),
          (reader.onerror = () =>
            reject(Error("Unable to read the generated template copy."))),
          reader.readAsArrayBuffer(blob));
      });
}
function describeImportSource(workbook, extension) {
  const baseName =
      workbook.name.replace(/[^a-zA-Z0-9 _-]/g, "").slice(0, 80) || "Workbook",
    sourceFileNames = [
      workbook.xlsxSource?.fileName,
      workbook.alternateSource?.fileName,
    ].map((fileName) => fileName?.toLowerCase());
  let derivedName = baseName + "-derived";
  for (
    ;
    sourceFileNames.includes((derivedName + "." + extension).toLowerCase());
  )
    derivedName += "-copy";
  return derivedName + "." + extension;
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
