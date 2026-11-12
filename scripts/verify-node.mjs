/**
 * Checks the built Node entry (lib/node.js) the way a script would use it: no browser, no
 * bundler. Run `npm run build:lib` first.
 */
import assert from "node:assert/strict";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import * as ledgerglass from "../lib/node.js";

const {
  FileProcessingError,
  addWorksheet,
  calculateWorkbookValues,
  createBlankWorkbook,
  createWorksheetTable,
  readWorkbookFile,
  writeWorkbookFile,
} = ledgerglass;

const directory = await mkdtemp(join(tmpdir(), "ledgerglass-node-"));
try {
  const workbook = createBlankWorkbook("Budget");
  Object.assign(workbook.sheets[0].cells, {
    A1: { raw: "Item" },
    B1: { raw: "Amount" },
    A2: { raw: "Rent" },
    B2: { raw: 1200 },
    A3: { raw: "Food" },
    B3: { raw: 450.5 },
    B4: { raw: "=SUM(B2:B3)" },
  });
  assert.deepEqual(calculateWorkbookValues(workbook).sheets.Sheet1.B4, {
    type: "number",
    value: 1650.5,
  });
  console.log("PASS the model calculates in Node");

  for (const extension of ["xlsx", "json", "csv", "tsv"]) {
    const path = join(directory, `budget.${extension}`);
    const written = await writeWorkbookFile(workbook, path);
    assert.ok(written.bytes > 0);
    const read = await readWorkbookFile(path);
    const values = Object.values(
      calculateWorkbookValues(read.workbook).sheets,
    )[0];
    assert.deepEqual(values.B4, { type: "number", value: 1650.5 }, extension);
    assert.equal(values.A2.value, "Rent", extension);
  }
  console.log("PASS XLSX, native JSON, CSV and TSV round-trip through files");

  await assert.rejects(
    readWorkbookFile(join(directory, "budget.json").replace("budget", "nope")),
    { code: "ENOENT" },
  );
  await assert.rejects(
    writeWorkbookFile(workbook, join(directory, "budget.pdf")),
    TypeError,
  );
  const { writeFile } = await import("node:fs/promises");
  await writeFile(join(directory, "broken.xlsx"), "not a zip");
  await assert.rejects(
    readWorkbookFile(join(directory, "broken.xlsx")),
    (error) =>
      error instanceof FileProcessingError && error.code === "FILE_INVALID",
  );
  console.log("PASS unreadable, unsupported and invalid files are refused");

  assert.equal(addWorksheet(workbook, "Notes").sheets.length, 2);
  assert.equal(typeof createWorksheetTable, "function");
  for (const name of [
    "AuditWorkbook",
    "AgentWorkbook",
    "createWorkbookCatalog",
  ])
    assert.equal(
      name in ledgerglass,
      false,
      `${name} must not be in the Node entry`,
    );
  console.log(
    `PASS the wider API is exported (${Object.keys(ledgerglass).length} names) without the interface`,
  );
} finally {
  await rm(directory, { recursive: true, force: true });
}
