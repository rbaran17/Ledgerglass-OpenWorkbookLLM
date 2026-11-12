// Imports a semicolon-separated CSV with decimal commas, adds a total row,
// saves it as XLSX, and reads the XLSX back to check the total.
//
// Plain delimited import would keep "1200,50" as text, because "auto" only
// reads numbers with a decimal point. So the Amount column (index 1, zero-based)
// is typed as a number, and "," is named as the decimal separator.
//
// Run from the repository root:
//   node examples/node-scripts/03-csv-to-xlsx.mjs [output-folder]
// With no folder, files go to a temporary folder that is removed at the end.

import { mkdtemp, mkdir, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import {
  applyCellPatchTransaction,
  calculateWorkbookValues,
  readWorkbookFile,
  writeWorkbookFile,
} from "@ledgerglass/workbook/node";

const givenFolder = process.argv[2];
const folder =
  givenFolder ?? (await mkdtemp(join(tmpdir(), "ledgerglass-03-")));
await mkdir(folder, { recursive: true });

// The CSV is written here so the example is self-contained.
const csvPath = join(folder, "expenses.csv");
await writeFile(
  csvPath,
  ["Item;Amount", "Rent;1200,50", "Food;450,25", "Travel;300,5"].join("\n") +
    "\n",
  "utf8",
);

// The delimited options are the ones from docs/api/files.md.
const { workbook: imported } = await readWorkbookFile(csvPath, {
  delimited: { delimiter: ";", decimal: ",", columns: { 1: "number" } },
});
const sheet = imported.sheets[0];
console.log(`Imported sheet: ${sheet.name}`);
console.log(`Revision before total: ${imported.revision}`);
console.log(`Events before total: ${imported.events.length}`);

// Append a total row under the three data rows (rows 2 to 4).
// This is a cell edit, so applyCellPatchTransaction records it as an event.
const withTotal = applyCellPatchTransaction(
  imported,
  [
    { sheetId: sheet.id, address: "A5", cell: { raw: "Total" } },
    { sheetId: sheet.id, address: "B5", cell: { raw: "=SUM(B2:B4)" } },
  ],
  "Example",
  "Append total row",
);
console.log(`Revision after total: ${withTotal.revision}`);
console.log(`Events after total: ${withTotal.events.length}`);

const xlsxPath = join(folder, "expenses.xlsx");
await writeWorkbookFile(withTotal, xlsxPath);

// Read the XLSX back and calculate it, to confirm the saved file carries the total.
const { workbook: reread } = await readWorkbookFile(xlsxPath);
const { sheets } = calculateWorkbookValues(reread);
const total = sheets[sheet.name].B5;
console.log(`Total read back from XLSX: ${total.value}`);

if (!givenFolder) await rm(folder, { recursive: true, force: true });
