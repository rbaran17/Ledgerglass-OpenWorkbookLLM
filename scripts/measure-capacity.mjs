/** Reproducible release-capacity measurements against the compiled public SDK. */
import { performance } from "node:perf_hooks";
import { cpus, platform, release, totalmem } from "node:os";
import { readFile, writeFile } from "node:fs/promises";
import { createHash } from "node:crypto";
import assert from "node:assert/strict";
import {
  createBlankWorkbook,
  createCalculator,
  applyPatches,
  validateWorkbook,
  importFile,
  exportCsv,
  exportXlsx,
  serializeWorkbook,
  deserializeWorkbook,
  WORKBOOK_CAPACITY,
} from "../lib/ledgerglass.js";
const out = new URL("../docs/CAPACITY-BENCHMARK.json", import.meta.url);
const samples = 5,
  measurements = [];
const round = (n) => Math.round(n * 100) / 100;
async function measure(name, operation, check) {
  const durations = [];
  for (let i = 0; i < samples; i++) {
    const start = performance.now();
    const result = await operation(i);
    durations.push(performance.now() - start);
    check?.(result);
  }
  const sorted = [...durations].sort((a, b) => a - b);
  measurements.push({
    name,
    samples: durations.map(round),
    medianMs: round(sorted[2]),
    p95Ms: round(sorted[Math.ceil(samples * 0.95) - 1]),
  });
}
const csv = Array.from({ length: 20000 }, (_, r) =>
  Array.from({ length: 5 }, (_, c) => String((r + 1) * 10 + c)).join(","),
).join("\r\n");
const file = new File([csv], "capacity-100000.csv");
const book = (await importFile(file)).workbook,
  sheet = book.sheets[0];
assert.equal(Object.keys(sheet.cells).length, 100000);
assert.equal(sheet.cells.E20000.raw, 200004);
const native = serializeWorkbook(book),
  xlsx = await exportXlsx(book);
await measure(
  "CSV import: 20,000 rows × 5 columns / 100,000 cells",
  () => importFile(file),
  (r) => assert.equal(Object.keys(r.workbook.sheets[0].cells).length, 100000),
);
await measure(
  "Validate 100,000-cell workbook",
  () => validateWorkbook(book),
  (r) => assert.equal(r.sheets[0].cells.E20000.raw, 200004),
);
await measure(
  "One cell patch on 100,000 cells (core transaction; excludes React/paint)",
  (i) =>
    applyPatches(book, [
      { sheetId: sheet.id, address: "A1", cell: { raw: 100 + i } },
    ]),
  (r) => assert.equal(r.sheets[0].cells.E20000.raw, 200004),
);
const expected = ((20000 * 20001) / 2) * 10;
await measure(
  "SUM of 20,000 numeric cells using a fresh calculator",
  () =>
    createCalculator(book).evaluateFormula(sheet.id, "F1", "=SUM(A1:A20000)"),
  (r) => assert.equal(r, expected),
);
await measure(
  "Native serialize 100,000 cells",
  () => serializeWorkbook(book),
  (r) => assert.equal(r, native),
);
await measure(
  "Native reopen 100,000 cells",
  () => deserializeWorkbook(native),
  (r) => assert.equal(r.sheets[0].cells.E20000.raw, 200004),
);
await measure(
  "CSV export 100,000-cell rectangle",
  () => exportCsv(book, sheet.id),
  (r) => assert.equal(r, csv),
);
await measure(
  "XLSX import 100,000 stored cells",
  () => importFile(new File([xlsx], "capacity.xlsx")),
  (r) => assert.equal(Object.keys(r.workbook.sheets[0].cells).length, 100000),
);
const sparse = createBlankWorkbook("Sparse far-corner formatting");
sparse.sheets[0].cells.A1 = { raw: 7 };
sparse.sheets[0].cells.XFD1048576 = {
  raw: null,
  style: { bold: true },
  note: "Far corner",
};
const sparseNative = serializeWorkbook(sparse),
  sparseXlsx = await exportXlsx(sparse);
await measure(
  "Sparse A1 + formatted XFD1048576 XLSX import",
  () => importFile(new File([sparseXlsx], "sparse.xlsx")),
  (r) => {
    assert.equal(Object.keys(r.workbook.sheets[0].cells).length, 2);
    assert.equal(r.workbook.sheets[0].cells.XFD1048576.style.bold, true);
  },
);
const outside = new File(
  [
    Array.from({ length: 100000 }, () =>
      Array.from({ length: 20 }, () => 1).join(","),
    ).join("\n"),
  ],
  "excluded-2million.csv",
);
const overStart = performance.now();
let refusal;
try {
  await importFile(outside);
  throw new Error("Excluded 2-million-cell fixture unexpectedly accepted");
} catch (error) {
  refusal = String(error.message);
  assert.match(refusal, /100,000/);
}
const result = {
  schemaVersion: 1,
  requirement: "AW-010",
  artifacts: Object.fromEntries(
    await Promise.all(
      [
        "lib/ledgerglass.js",
        "lib/ledgerglass.css",
        "scripts/measure-capacity.mjs",
      ].map(async (path) => [
        path,
        createHash("sha256")
          .update(await readFile(new URL("../" + path, import.meta.url)))
          .digest("hex"),
      ]),
    ),
  ),
  measuredAt: new Date().toISOString(),
  environment: {
    runtime: process.version,
    platform: platform(),
    osRelease: release(),
    architecture: process.arch,
    cpu: cpus()[0]?.model,
    logicalCpus: cpus().length,
    systemMemoryBytes: totalmem(),
    processMaxRssBytes: process.resourceUsage().maxRSS * 1024,
  },
  method: {
    samplesPerOperation: samples,
    warmup:
      "Fixture preparation imports once; each listed operation then runs five times sequentially.",
    percentile:
      "Nearest rank; p95 with five samples is the maximum, not a statistically stable performance guarantee.",
    process:
      "Shared development machine; measurements include synchronous processing on this Node process. Memory is the process high-water RSS, not incremental workbook allocation.",
    ui: "Node timings exclude React and painting. See CAPACITY-BROWSER-EVIDENCE.md for the separately observed production browser run.",
  },
  capacity: WORKBOOK_CAPACITY,
  fixtures: {
    dense: {
      rows: 20000,
      columns: 5,
      storedCells: 100000,
      csvBytes: file.size,
      nativeBytes: Buffer.byteLength(native),
      xlsxBytes: xlsx.size,
    },
    sparse: {
      lastAddress: "XFD1048576",
      storedCells: 2,
      nativeBytes: Buffer.byteLength(sparseNative),
      xlsxBytes: sparseXlsx.size,
    },
    excluded: {
      rows: 100000,
      columns: 20,
      cells: 2000000,
      csvBytes: outside.size,
      refusal,
      refusalMs: round(performance.now() - overStart),
    },
  },
  measurements,
  acceptance: {
    capacityContractVerified: true,
    performanceSlaAccepted: false,
    AW102:
      "Partial; 2,000,000-cell populations are outside this release capacity. Worker processing and cross-device/browser performance acceptance remain open.",
    externalLLMTested: false,
  },
};
await writeFile(out, JSON.stringify(result, null, 2) + "\n");
console.log(
  JSON.stringify(
    {
      report: out.pathname,
      measurements,
      fixtureBytes: result.fixtures,
      processMaxRssBytes: result.environment.processMaxRssBytes,
    },
    null,
    2,
  ),
);
