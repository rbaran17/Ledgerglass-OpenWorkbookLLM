// Creates a worksheet table over a range, adds a data validation rule to one
// of its columns, and checks a valid and an invalid value against the rule.
//
// The table is added with a function from docs/api/tables.md, which records an event.
// The validation rule is stored directly on the sheet, as docs/api/rules.md shows: the
// SDK has no function that adds a rule, so that write is not recorded in the event log.
//
// Run from the repository root:
//   node examples/node-scripts/04-tables-and-validation.mjs

import {
  DataValidationError,
  applyCellPatchTransaction,
  checkDataValidationBatch,
  createBlankWorkbook,
  createWorksheetTable,
  normalizeDataValidationRules,
  validateDataValidationRule,
} from "@ledgerglass/workbook/node";

// Sales data: a header row and three rows of numbers.
let workbook = createBlankWorkbook("Sales");
const sheet = workbook.sheets[0];
workbook = applyCellPatchTransaction(
  workbook,
  [
    { sheetId: sheet.id, address: "A1", cell: { raw: "Region" } },
    { sheetId: sheet.id, address: "B1", cell: { raw: "Units" } },
    { sheetId: sheet.id, address: "C1", cell: { raw: "Price" } },
    { sheetId: sheet.id, address: "A2", cell: { raw: "North" } },
    { sheetId: sheet.id, address: "B2", cell: { raw: 10 } },
    { sheetId: sheet.id, address: "C2", cell: { raw: 2.5 } },
    { sheetId: sheet.id, address: "A3", cell: { raw: "South" } },
    { sheetId: sheet.id, address: "B3", cell: { raw: 4 } },
    { sheetId: sheet.id, address: "C3", cell: { raw: 3 } },
    { sheetId: sheet.id, address: "A4", cell: { raw: "West" } },
    { sheetId: sheet.id, address: "B4", cell: { raw: 7 } },
    { sheetId: sheet.id, address: "C4", cell: { raw: 1.5 } },
  ],
  "Example",
  "Enter sales",
);

// A table gives the range a header row and named columns. The first row of the
// range supplies the column names.
workbook = createWorksheetTable(workbook, sheet.id, {
  name: "Sales",
  range: "A1:C4",
});

// Print a short summary of the table.
const table = workbook.sheets[0].tables[0];
console.log(`Table: ${table.name}`);
console.log(`Range: ${table.range}`);
console.log(
  `Columns: ${table.columns.map((column) => column.name).join(", ")}`,
);

// Validation rule: Units in B2:B4 must be a whole-ish number from 0 to 100.
// normalizeDataValidationRules checks the rule and returns a clean copy.
// The rule is stored on the sheet, as docs/api/rules.md shows.
const [rule] = normalizeDataValidationRules([
  {
    id: "units",
    range: "B2:B4",
    kind: "number",
    operator: "between",
    min: 0,
    max: 100,
    list: [],
    allowBlank: false,
    message: "Units must be between 0 and 100.",
  },
]);
workbook.sheets[0].validations = [rule];

// Check values against the rule directly. A passing value returns true.
console.log(`Rule accepts 7: ${validateDataValidationRule(rule, 7)}`);
console.log(`Rule accepts 150: ${validateDataValidationRule(rule, 150)}`);

// Check cells in the workbook. checkDataValidationBatch returns 0 when every
// target passes, and throws DataValidationError when one does not.
const targets = [{ sheetId: sheet.id, address: "B2" }];
console.log(
  `Batch check of B2 (7): ${checkDataValidationBatch(workbook, targets)} problems`,
);

// Put an invalid value in B3 and check it. applyCellPatchTransaction would
// refuse this value (the rule is a stop rule), so the copy is made by hand
// here: a new sheet object with one more cell, validations kept.
const current = workbook.sheets[0]; // the sheet that now holds the rule
const badSheet = { ...current, cells: { ...current.cells, B3: { raw: 150 } } };
const invalid = { ...workbook, sheets: [badSheet] };
try {
  checkDataValidationBatch(invalid, [{ sheetId: sheet.id, address: "B3" }]);
  console.log("Batch check of B3 (150): passed");
} catch (error) {
  if (error instanceof DataValidationError) {
    // The error lists each failed check; here there is one.
    console.log(`Batch check of B3 (150): ${error.issues.length} problem(s)`);
    console.log(`Message: ${error.message}`);
  } else {
    throw error;
  }
}
