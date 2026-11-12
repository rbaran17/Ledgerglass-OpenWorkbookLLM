// Builds a workbook from a description an AI agent sent, prints the summary the
// agent would get back, then shows a description that is rejected and its problems.
//
// The description format is the one in docs/AGENT-UI.md, section 1. Treat any
// description as untrusted input: buildWorkbookFromSpec checks it against the
// limits and throws WorkbookSpecError, which lists every problem in plain words.
//
// Run from the repository root:
//   node examples/node-scripts/05-agent-workbook-check.mjs

import {
  WorkbookSpecError,
  buildWorkbookFromSpec,
  summarizeWorkbookForAgent,
} from "@ledgerglass/workbook/node";

// A description that is accepted: one sheet, a header row, and a formula
// that divides by zero so the summary has an error to report.
const goodSpec = {
  title: "Budget",
  sheets: [
    {
      name: "Plan",
      rows: [
        ["Item", "Cost"],
        ["Rent", 1200],
        ["Food", 450],
      ],
      cells: { B4: "=SUM(B2:B3)", C4: "=1/0" },
    },
  ],
};

// The id keeps the same document when the same description is shown again.
const { workbook, title, readOnly } = buildWorkbookFromSpec(goodSpec, {
  id: "budget-demo",
});
console.log(`Title: ${title}`);
console.log(`Read only: ${readOnly}`);
console.log(`Sheets: ${workbook.sheets.map((sheet) => sheet.name).join(", ")}`);

// This is what the agent receives back: counts, and the formulas that failed.
const summary = summarizeWorkbookForAgent(workbook);
console.log("Summary for the agent:");
console.log(JSON.stringify(summary, null, 2));

// A description that breaks three rules: a sheet name with "/", a duplicate
// name that differs only in case, and an empty sheet name.
const badSpec = {
  title: "Broken",
  sheets: [
    { name: "Plan/2026", rows: [["Item"]] },
    { name: "Costs", rows: [["Item"]] },
    { name: "costs", rows: [["Item"]] },
    { name: "", rows: [["Item"]] },
  ],
};

try {
  buildWorkbookFromSpec(badSpec);
  console.log("Unexpected: the description was accepted.");
} catch (error) {
  if (!(error instanceof WorkbookSpecError)) throw error;
  // problems is the list to show the user or to send back to the agent.
  console.log(`Rejected with ${error.problems.length} problem(s):`);
  for (const problem of error.problems) console.log(`- ${problem}`);
}
