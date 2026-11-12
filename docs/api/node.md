# Node entry

`@ledgerglass/workbook/node` is the library without the interface: the workbook model, calculation, files, the assistant protocol and the agent workbook description. It runs in Node 22 or later with no browser and no bundler. Use it in scripts, servers and tests.

It exports everything on the other [API reference](README.md) pages, the core functions (`createBlankWorkbook`, `normalizeWorkbookInput`, `getWorkbookCalculation`, `applyCellPatchTransaction`, `buildCanonicalWorkbookSnapshot`, `shiftFormulaReferences`, `workbookLimits`), the assistant protocol functions, the workbook description functions from [the agent guide](../AGENT-UI.md), and the three helpers below. It does not export the editor component, the assistant adapters, `AgentWorkbook` or `createWorkbookCatalog`.

Build it with `npm run build:lib`, which writes `lib/node.js`. `npm run test:lib` checks it in plain Node.

```js
import {
  calculateWorkbookValues,
  readWorkbookFile,
  writeWorkbookFile,
} from "@ledgerglass/workbook/node";

const { workbook, warnings } = await readWorkbookFile("budget.xlsx");
const { sheets } = calculateWorkbookValues(workbook);
console.log(sheets.Budget.B4); // { type: "number", value: 1650.5 }
await writeWorkbookFile(workbook, "budget.json"); // native JSON keeps everything
```

## readWorkbookFile

```ts
function readWorkbookFile(
  path: string,
  options?: {
    delimited?: DelimitedImportOptions;
    signal?: AbortSignal;
    timeoutMs?: number;
  },
): Promise<{
  workbook: Workbook;
  warnings: string[];
  report: object | undefined;
}>;
```

Reads a workbook from a path. The extension picks the format:

| Extension              | Read as                                                        |
| ---------------------- | -------------------------------------------------------------- |
| `.json`                | Native JSON, validated with `normalizeWorkbookInput`           |
| `.csv`, `.tsv`, `.txt` | Delimited text; `options.delimited` sets the parser options    |
| anything else          | The workbook importer (`importWorkbookFile`), which reads XLSX |

`warnings` lists what an XLSX import could not carry over; it is empty for the other formats. `report` is the importer's report and is `undefined` for native JSON. Parsing runs in a worker thread.

It rejects with the file system's error (for example `ENOENT`) if the path cannot be read, with a [`FileProcessingError`](files.md#errors) if the file is not accepted, and with an `Error` if a `.json` file is not a workbook.

## writeWorkbookFile

```ts
function writeWorkbookFile(
  workbook: Workbook,
  path: string,
  options?: { sheetId?: string; delimited?: DelimitedExportOptions },
): Promise<{ path: string; bytes: number; warnings: string[] }>;
```

Writes a workbook to a path. The extension picks the format: `.xlsx`, `.json` (native JSON, which keeps everything) or `.csv` / `.tsv` (one sheet, with formulas written as their calculated values). For delimited text, `options.sheetId` chooses the sheet (the first by default) and `options.delimited` passes options to [`exportWorksheetToCsv`](files.md#exportworksheettocsv); `.tsv` uses a tab unless you set a delimiter.

`warnings` lists what the format could not keep; for XLSX it always includes the general notes described under [`exportWorkbookToXlsx`](files.md#exportworkbooktoxlsx). Any other extension throws a `TypeError`.

## calculateWorkbookValues

```ts
function calculateWorkbookValues(workbook: Workbook): {
  report: CalculationReport;
  sheets: Record<string, Record<string, TypedValue>>;
};
```

Calculates the workbook and returns the result of every filled cell as `sheets[sheetName][address]`. Each result is `{ type, value }`, where `type` is `"number"`, `"text"`, `"boolean"`, `"error"` or `"blank"`; a formula error appears as `{ type: "error", value: "#DIV/0!" }`. `report` is the calculation report. For single cells or formulas that are not stored in the workbook, use the calculator from `getWorkbookCalculation` (see [Calculation](calculation.md)).

## Example: a command-line check

```js
import {
  readWorkbookFile,
  calculateWorkbookValues,
} from "@ledgerglass/workbook/node";

const { workbook } = await readWorkbookFile(process.argv[2]);
const { sheets } = calculateWorkbookValues(workbook);
const errors = Object.entries(sheets).flatMap(([sheet, cells]) =>
  Object.entries(cells)
    .filter(([, result]) => result.type === "error")
    .map(([address, result]) => `${sheet}!${address}: ${result.value}`),
);
console.log(errors.length ? errors.join("\n") : "No formula errors.");
process.exitCode = errors.length ? 1 : 0;
```
