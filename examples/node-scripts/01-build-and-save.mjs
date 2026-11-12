// Shows how to build a small workbook in code, calculate it, and save it.
//
// It builds a budget with a header row, numbers, a SUM and a percentage formula,
// adds a second sheet, freezes the header row, prints the calculated values, and
// writes the workbook as XLSX and as native JSON (the format that keeps everything).
//
// Run from the repository root:
//   node examples/node-scripts/01-build-and-save.mjs [output-folder]
// With no folder, files go to a temporary folder that is removed at the end.

import { mkdtemp, mkdir, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { basename, join } from "node:path";
import {
  addWorksheet,
  applyCellPatchTransaction,
  calculateWorkbookValues,
  createBlankWorkbook,
  renameWorksheet,
  setFreezePanes,
  writeWorkbookFile,
} from "@ledgerglass/workbook/node";

// Use the folder given on the command line, or a temporary one we remove later.
const givenFolder = process.argv[2];
const folder = givenFolder ?? (await mkdtemp(join(tmpdir(), "ledgerglass-01-")));
await mkdir(folder, { recursive: true });

// A new workbook has one sheet, "Sheet1". Rename it to "Budget".
let workbook = createBlankWorkbook("Budget");
workbook = renameWorksheet(workbook, workbook.sheets[0].id, "Budget");
const budget = workbook.sheets[0];

// Cell edits go through applyCellPatchTransaction. It is the only function that
// records an event and increases `revision`, so it is the right way to edit cells.
// Each patch sets one cell; a value starting with "=" is a formula.
workbook = applyCellPatchTransaction(
  workbook,
  [
    { sheetId: budget.id, address: "A1", cell: { raw: "Item" } },
    { sheetId: budget.id, address: "B1", cell: { raw: "Amount" } },
    { sheetId: budget.id, address: "A2", cell: { raw: "Rent" } },
    { sheetId: budget.id, address: "B2", cell: { raw: 1200 } },
    { sheetId: budget.id, address: "A3", cell: { raw: "Food" } },
    { sheetId: budget.id, address: "B3", cell: { raw: 450 } },
    { sheetId: budget.id, address: "A4", cell: { raw: "Travel" } },
    { sheetId: budget.id, address: "B4", cell: { raw: 300.5 } },
    { sheetId: budget.id, address: "A5", cell: { raw: "Total" } },
    { sheetId: budget.id, address: "B5", cell: { raw: "=SUM(B2:B4)" } },
    // Share of the total that rent takes; a fraction, shown as a decimal here.
    { sheetId: budget.id, address: "A6", cell: { raw: "Rent share" } },
    { sheetId: budget.id, address: "B6", cell: { raw: "=B2/B5" } },
  ],
  "Example",
  "Build budget",
);

// Adding a sheet returns a new workbook; the new sheet is last in the tab order.
workbook = addWorksheet(workbook, "Notes");
const notes = workbook.sheets[1];
workbook = applyCellPatchTransaction(
  workbook,
  [{ sheetId: notes.id, address: "A1", cell: { raw: "Figures are in euros." } }],
  "Example",
  "Add note",
);

// Keep the header row visible while scrolling. Freezing one row is a pane
// setting, so it does not add an event to the workbook.
workbook = setFreezePanes(workbook, budget.id, 1);

// calculateWorkbookValues returns the result of every filled cell, keyed by sheet name.
const { sheets } = calculateWorkbookValues(workbook);
console.log("Budget values:");
for (const address of ["B5", "B6"]) {
  console.log(`  ${address}: ${JSON.stringify(sheets.Budget[address])}`);
}

// Save the same workbook twice. XLSX is for Excel; native JSON keeps everything.
const xlsx = await writeWorkbookFile(workbook, join(folder, "budget.xlsx"));
const json = await writeWorkbookFile(workbook, join(folder, "budget.json"));
console.log(`Wrote ${basename(xlsx.path)}`);
console.log(`Wrote ${basename(json.path)}`);
// The XLSX warnings list what XLSX cannot keep (for example review history).
console.log(`XLSX warnings: ${xlsx.warnings.length}`);

if (!givenFolder) await rm(folder, { recursive: true, force: true });
