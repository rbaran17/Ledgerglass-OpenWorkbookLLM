# Node scripts

Five short scripts that use the Node entry, `@ledgerglass/workbook/node`, without a browser. They build and save workbooks, check them for formula errors, import CSV, work with tables and validation rules, and check a workbook description an AI agent sends. Each script runs from the repository root with `node`, needs no arguments (most accept an optional output folder), and cleans up any temporary files it creates. Each `NAME.expected.txt` holds the exact output of the script's default run.

## 01 build and save

Builds a small budget in code: a header row, numbers, a `SUM` and a percentage formula, a second sheet added with `addWorksheet`, and a frozen header row. It prints the calculated values with `calculateWorkbookValues`, then saves the workbook as `.xlsx` and as native `.json`, and prints the file names and the number of XLSX warnings.

```sh
node examples/node-scripts/01-build-and-save.mjs [output-folder]
```

```text
Budget values:
  B5: {"type":"number","value":1950.5}
  B6: {"type":"number","value":0.6152268649064343}
Wrote budget.xlsx
Wrote budget.json
XLSX warnings: 3
```

## 02 check formula errors

Reads a workbook file and lists every formula whose result is an error, as `Sheet!Address: error`. The exit code is 1 when there are errors and 0 otherwise. With no argument it writes its own small workbook with `=1/0` and `=#REF!` to a temporary folder and checks that.

```sh
node examples/node-scripts/02-check-formula-errors.mjs [path-to-workbook]
```

```text
Budget!B1: #DIV/0!
Budget!C1: #REF!
```

The default run exits with code 1. Given a workbook with no errors (for example the file that script 01 writes), it prints `No formula errors.` and exits with 0.

## 03 CSV to XLSX

Writes a semicolon-separated CSV with decimal commas, imports it with `readWorkbookFile` and the delimited options `{ delimiter: ";", decimal: ",", columns: { 1: "number" } }`, appends a total row with `applyCellPatchTransaction`, and shows that `revision` and `events` grew. It saves the result as `.xlsx`, reads that file back, and prints the total.

```sh
node examples/node-scripts/03-csv-to-xlsx.mjs [output-folder]
```

```text
Imported sheet: Imported data
Revision before total: 0
Events before total: 0
Revision after total: 1
Events after total: 1
Total read back from XLSX: 1951.25
```

## 04 tables and validation

Creates a worksheet table over `A1:C4` with `createWorksheetTable`, and a number rule on the Units column with `normalizeDataValidationRules`. It prints a short table summary, checks a valid value (7) and an invalid one (150) with `validateDataValidationRule`, and runs `checkDataValidationBatch` on a cell holding each value.

```sh
node examples/node-scripts/04-tables-and-validation.mjs
```

```text
Table: Sales
Range: A1:C4
Columns: Region, Units, Price
Rule accepts 7: true
Rule accepts 150: false
Batch check of B2 (7): 0 problems
Batch check of B3 (150): 1 problem(s)
Message: Sheet1!B3: Units must be between 0 and 100. Value does not satisfy the rule.
```

## 05 agent workbook check

Builds a workbook from a description in the format of `docs/AGENT-UI.md`, section 1, with `buildWorkbookFromSpec`. It prints the result of `summarizeWorkbookForAgent`, which is what the agent gets back, including a formula error. Then it builds a description that breaks three rules and prints the `problems` of the `WorkbookSpecError`.

```sh
node examples/node-scripts/05-agent-workbook-check.mjs
```

```text
Title: Budget
Read only: false
Sheets: Plan
Summary for the agent:
{
  "workbookId": "budget-demo",
  "title": "Budget",
  "sheets": [
    {
      "name": "Plan",
      "filledCells": 8,
      "formulaCount": 2
    }
  ],
  "formulaErrorCount": 1,
  "formulaErrors": [
    {
      "sheet": "Plan",
      "address": "C4",
      "formula": "=1/0",
      "error": "#DIV/0!"
    }
  ]
}
Rejected with 3 problem(s):
- Sheet 1: the name may not contain : \ / ? * [ ].
- Sheet 3: the name "costs" is used twice.
- Sheet 4 needs a name.
```
