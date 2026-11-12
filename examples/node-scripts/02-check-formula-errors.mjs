// Lists every formula error in a workbook file and sets the exit code.
//
// The file is read with readWorkbookFile (the extension picks the format), every
// formula is calculated with calculateWorkbookValues, and each cell whose result
// is an error (such as #DIV/0! or #REF!) is printed as "Sheet!Address: error".
// The exit code is 1 when there are errors, so this can run in a CI step.
//
// Run from the repository root:
//   node examples/node-scripts/02-check-formula-errors.mjs [path-to-workbook]
// With no path, the script first writes a small workbook with two bad formulas
// to a temporary folder, checks that, and removes the folder afterwards.

import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import {
  applyCellPatchTransaction,
  calculateWorkbookValues,
  createBlankWorkbook,
  readWorkbookFile,
  renameWorksheet,
  writeWorkbookFile,
} from "@ledgerglass/workbook/node";

// Turns "AA12" into { row: 12, col: 27 } so errors can be listed in reading order.
function position(address) {
  const [, letters, digits] = /^([A-Z]+)(\d+)$/.exec(address);
  let col = 0;
  for (const letter of letters) col = col * 26 + (letter.charCodeAt(0) - 64);
  return { row: Number(digits), col };
}

let path = process.argv[2];
let tempFolder;

if (!path) {
  // Build a demonstration file: one good formula and two that fail.
  tempFolder = await mkdtemp(join(tmpdir(), "ledgerglass-02-"));
  path = join(tempFolder, "with-errors.xlsx");
  let workbook = createBlankWorkbook("Errors");
  workbook = renameWorksheet(workbook, workbook.sheets[0].id, "Budget");
  const sheetId = workbook.sheets[0].id;
  workbook = applyCellPatchTransaction(
    workbook,
    [
      { sheetId, address: "A1", cell: { raw: 10 } },
      { sheetId, address: "B1", cell: { raw: "=1/0" } }, // division by zero
      { sheetId, address: "C1", cell: { raw: "=#REF!" } }, // a broken reference
      { sheetId, address: "D1", cell: { raw: "=A1*2" } }, // fine: 20
    ],
    "Example",
    "Add formulas",
  );
  await writeWorkbookFile(workbook, path);
}

// readWorkbookFile rejects if the file is missing or not accepted, so a bad path stops here.
const { workbook } = await readWorkbookFile(path);
const { sheets } = calculateWorkbookValues(workbook);

// Walk sheets in tab order and cells in reading order, so the output is stable.
const errors = [];
for (const sheet of workbook.sheets) {
  const cells = sheets[sheet.name] ?? {};
  const addresses = Object.keys(cells).sort((a, b) => {
    const pa = position(a);
    const pb = position(b);
    return pa.row - pb.row || pa.col - pb.col;
  });
  for (const address of addresses) {
    const result = cells[address];
    if (result.type === "error") {
      errors.push(`${sheet.name}!${address}: ${result.value}`);
    }
  }
}

console.log(errors.length ? errors.join("\n") : "No formula errors.");
process.exitCode = errors.length ? 1 : 0;

if (tempFolder) await rm(tempFolder, { recursive: true, force: true });
