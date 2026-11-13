# Calculation

The calculation API sets how a workbook recalculates, runs calculations with row filters, checks the evidence a calculation leaves behind, and lists the function names the calculator knows. It also validates what-if analyses and reports the resource budgets a calculation runs under. Use `getWorkbookCalculation` from the workbook API to read values; the calculator it returns is described in [The calculator object](#the-calculator-object).

Types are in `types/sdk/calculation.d.ts`. None of these functions changes the cells or the revision of the workbook it receives. Functions that return a workbook return a new object.

## Settings

### normalizeCalculationSettings

```ts
export function normalizeCalculationSettings(
  settings: unknown,
): CalculationSettings;
```

Returns a normalized copy of the settings. Throws an `Error` whose message starts with "Invalid calculation settings." when a key is unknown, a value is out of range, or the argument is not an object. Every key except `memoryBudgetMiB` and `workBudget` is required.

```js
import {
  defaultCalculationSettings,
  normalizeCalculationSettings,
} from "@ledgerglass/workbook";

const settings = normalizeCalculationSettings({
  ...defaultCalculationSettings,
  mode: "manual",
});
```

### defaultCalculationSettings

```ts
export const defaultCalculationSettings: Readonly<CalculationSettings>;
```

The defaults: `mode` "automatic", `iterative` false, `maxIterations` 100, `maxChange` 0.001 and `timeBudgetMs` 2500. The object is frozen. Spread it to change one setting, then pass the result to `normalizeCalculationSettings` or `withCalculationSettings`.

### withCalculationSettings

```ts
export function withCalculationSettings(
  workbook: Workbook,
  settings: CalculationSettings,
): Workbook;
```

Returns a new workbook object whose `calculation` field holds the normalized settings. The new object shares its `sheets` array with the argument, and the argument is not changed. The returned workbook's `revision` is one higher and its `events` gets one new event, `Change calculation settings`, with an empty change list. Throws when the settings are invalid.

```js
import {
  createBlankWorkbook,
  defaultCalculationSettings,
  withCalculationSettings,
} from "@ledgerglass/workbook";

const workbook = createBlankWorkbook("Budget");
const manual = withCalculationSettings(workbook, {
  ...defaultCalculationSettings,
  mode: "manual",
});
```

## Running calculations

### calculateWorkbookWithSettings

```ts
export function calculateWorkbookWithSettings(
  workbook: Workbook,
  options?: CalculationOptions,
): CalculatorMethods;
```

Calculates the workbook with row filters applied and returns the calculator. The only accepted option is `filteredRowsBySheet`, a map from worksheet ID to zero-based, inclusive `{ start, end }` intervals, and `SUBTOTAL` skips the filtered rows. Other keys throw. A plain workbook's result is not stored.

```js
import {
  calculateWorkbookWithSettings,
  createBlankWorkbook,
} from "@ledgerglass/workbook";

const workbook = createBlankWorkbook("Budget");
const sheet = workbook.sheets[0];
sheet.cells.A1 = { raw: 120 };
sheet.cells.A2 = { raw: 80 };
sheet.cells.A3 = { raw: 50 };
sheet.cells.B1 = { raw: "=SUBTOTAL(9,A1:A3)" };

const calculator = calculateWorkbookWithSettings(workbook, {
  filteredRowsBySheet: { [sheet.id]: [{ start: 1, end: 1 }] },
});
calculator.getTypedValue(sheet.id, "B1"); // { type: "number", value: 170 }
```

### refreshCalculationReport

```ts
export function refreshCalculationReport(workbook: Workbook): CalculationResult;
```

Recalculates every formula now, even when a current result exists, and stores the new result. Use it for volatile functions such as `RAND`, and to turn a pending manual-mode result into a value. Returns `{ calculator, report }`. The report's `run` count increases by one.

```js
import {
  createBlankWorkbook,
  defaultCalculationSettings,
  getWorkbookCalculation,
  refreshCalculationReport,
  withCalculationSettings,
} from "@ledgerglass/workbook";

const workbook = createBlankWorkbook("Budget");
workbook.sheets[0].cells.A1 = { raw: 2 };
workbook.sheets[0].cells.B1 = { raw: "=A1*3" };
const manual = withCalculationSettings(workbook, {
  ...defaultCalculationSettings,
  mode: "manual",
});

getWorkbookCalculation(manual).calculator.getTypedValue(
  manual.sheets[0].id,
  "B1",
);
// { type: "error", value: "#PENDING!" }

const { calculator } = refreshCalculationReport(manual);
calculator.getTypedValue(manual.sheets[0].id, "B1"); // { type: "number", value: 6 }
```

### getCurrentCalculatorRecord

```ts
export function getCurrentCalculatorRecord(
  workbook: Workbook,
): CalculationResult | undefined;
```

Returns the stored `{ calculator, report }` for this exact workbook object, or undefined when no result is current. It never calculates. A new object built from an edited workbook has no stored result, so the function returns undefined for it.

```js
import {
  createBlankWorkbook,
  getCurrentCalculatorRecord,
  getWorkbookCalculation,
} from "@ledgerglass/workbook";

const workbook = createBlankWorkbook("Budget");
getCurrentCalculatorRecord(workbook); // undefined
getWorkbookCalculation(workbook);
getCurrentCalculatorRecord(workbook).report; // the report from getWorkbookCalculation
```

### buildCalculationEvidence

```ts
export function buildCalculationEvidence(
  workbook: Workbook,
  report: CalculationReport,
): string;
```

Returns a JSON string with `schemaVersion`, the report, the imported calculation caches (or `null`) and a list of boundaries that state what the evidence does not prove. Throws "The workbook or calculation run changed." when `report` is not the current report for `workbook`, so call it with the report from the latest calculation.

```js
import {
  buildCalculationEvidence,
  createBlankWorkbook,
  getWorkbookCalculation,
} from "@ledgerglass/workbook";

const workbook = createBlankWorkbook("Budget");
const { report } = getWorkbookCalculation(workbook);
const evidence = JSON.parse(buildCalculationEvidence(workbook, report));
evidence.schemaVersion; // 1
```

## Function names

Each of these is a read-only list. The arrays are shared with the calculator and are not frozen, so do not change them.

### builtinFunctionNames

```ts
export const builtinFunctionNames: ReadonlyArray<string>;
```

Every built-in worksheet function name the calculator evaluates, in uppercase. Contains about 395 names, including `SUM` and `XLOOKUP`.

### statisticalFunctionNames

```ts
export const statisticalFunctionNames: ReadonlyArray<string>;
```

The statistical functions, such as `AVERAGE` and `MEDIAN`. Every entry is also in `builtinFunctionNames`.

### databaseFunctionNames

```ts
export const databaseFunctionNames: ReadonlyArray<string>;
```

The database functions, such as `DSUM` and `DGET`. There are 12.

### matrixResultFunctionNames

```ts
export const matrixResultFunctionNames: ReadonlyArray<string>;
```

Functions that return a matrix, such as `SEQUENCE` and `FILTER`. `SUM` is not in the list.

### daxFunctionNames

```ts
export const daxFunctionNames: Readonly<
  Record<string, readonly [number, number]>
>;
```

The DAX functions the data model accepts. Each value is `[minimum, maximum]` argument counts, so `IF` is `[2, 3]` and `BLANK` is `[0, 0]`. The object is frozen.

### textCleanupFunctionNames

```ts
export const textCleanupFunctionNames: ReadonlyArray<string>;
```

Text functions in the cleanup group, such as `TRIM` and `CLEAN`.

### dateFunctionNames

```ts
export const dateFunctionNames: ReadonlyArray<string>;
```

Date functions, such as `DATE` and `EDATE`. There are 25.

### supportedFinancialFunctionNames

```ts
export const supportedFinancialFunctionNames: ReadonlyArray<string>;
```

Financial functions the calculator supports, such as `PMT` and `NPV`. There are 20.

### supportedTextFunctionNames

```ts
export const supportedTextFunctionNames: ReadonlyArray<string>;
```

Text functions the calculator supports, such as `LEFT` and `TEXTJOIN`. There are 22.

## Budgets and limits

### resolveCalculationBudget

```ts
export function resolveCalculationBudget(calculation?: {
  memoryBudgetMiB?: number;
  workBudget?: number;
}): CalculationBudget;
```

Converts calculation settings into `memoryBudgetBytes` and `workBudget`. A missing memory budget means 32 MiB and a missing work budget means 1,000,000 units. The function does not validate its input, so pass normalized settings. The calculator's resource report uses the same values.

```js
import { resolveCalculationBudget } from "@ledgerglass/workbook";

resolveCalculationBudget({ memoryBudgetMiB: 16 });
// { memoryBudgetBytes: 16777216, workBudget: 1000000 }
```

### estimateWorkbookMemoryBytes

```ts
export function estimateWorkbookMemoryBytes(
  workbook: Workbook,
  memoryBudgetBytes: number,
): number;
```

Estimates the working data the workbook needs, in bytes. Counting stops once the total passes `memoryBudgetBytes`, so a result above the budget is not the exact total. Throws a `TypeError` when the argument has no `sheets` array.

```js
import {
  createBlankWorkbook,
  estimateWorkbookMemoryBytes,
} from "@ledgerglass/workbook";

const workbook = createBlankWorkbook("Budget");
estimateWorkbookMemoryBytes(workbook, 33554432); // 1324 for a blank workbook
```

### worksheetCapacityLimits

```ts
export const worksheetCapacityLimits: Readonly<{
  maxIntervalsPerAxis: number;
  maxFrozenVisibleRows: number;
  maxFrozenVisibleCols: number;
  maxXlsxHiddenRows: number;
}>;
```

Fixed capacity limits: 10,000 hidden intervals per axis, 100 frozen visible rows, 50 frozen visible columns and 100,000 hidden rows in an XLSX export.

### whatIfLimits

```ts
export const whatIfLimits: Readonly<{
  inputs: number;
  outputs: number;
  scenarios: number;
  axisValues: number;
  trials: number;
  dependencyCells: number;
  workUnits: number;
  savedAnalyses: number;
}>;
```

Fixed what-if limits: 32 inputs, 16 outputs, 100 scenarios, 100 axis values, 400 trials, 500 dependency cells, 50,000 work units and 40 saved analyses.

## What-if analyses

### validateWhatIfDefinition

```ts
export function validateWhatIfDefinition(request: unknown): WhatIfDefinition;
```

Checks a what-if request and returns a normalized definition with trimmed names and A1 addresses. The request has a `kind` of `"scenarios"`, `"one-variable"` or `"two-variable"`. Throws with a message that says which rule failed, such as duplicate scenario names (compared without case), an input that is also an output, or a two-variable table with more than 400 combinations.

```js
import { validateWhatIfDefinition } from "@ledgerglass/workbook";

const definition = validateWhatIfDefinition({
  name: "Rates",
  kind: "one-variable",
  inputs: [{ sheetId: "sheet-1", address: "A1" }],
  outputs: [{ sheetId: "sheet-1", address: "B2" }],
  orientation: "row",
  values: [100, 150],
});
```

### expandWhatIfScenarios

```ts
export function expandWhatIfScenarios(request: WhatIfDefinition): WhatIfCase[];
```

Expands a validated definition into its cases, each with a `name` and the `inputs` to calculate. Scenarios use their own names. One-variable cases are named by their value. Two-variable cases are named `"row / column"`. The function does not validate its argument, so an unreadable request throws a `TypeError`.

```js
import {
  expandWhatIfScenarios,
  validateWhatIfDefinition,
} from "@ledgerglass/workbook";

const definition = validateWhatIfDefinition({
  name: "Rates",
  kind: "one-variable",
  inputs: [{ sheetId: "sheet-1", address: "A1" }],
  outputs: [{ sheetId: "sheet-1", address: "B2" }],
  orientation: "row",
  values: [100, 150],
});
expandWhatIfScenarios(definition);
// [{ name: "100", inputs: [100] }, { name: "150", inputs: [150] }]
```

## The calculator object

`getWorkbookCalculation(workbook).calculator` is a function with 13 own methods. Calling it as `calculator(sheetId, address)` returns the cell's value, the same as `calculator.getTypedValue(sheetId, address).value`. `Object.keys` lists the 13 methods, and the function's prototype is `Function.prototype`. `calculateWorkbookWithSettings` returns the calculator directly. `refreshCalculationReport` and `getWorkbookCalculation` return it as `calculator` next to the `report`. The shape is typed as `CalculatorMethods` in `types/sdk/calculation.d.ts`.

Sheet IDs are the `id` of a sheet. Addresses are A1-style, case-insensitive, and may contain `$`.

### getTypedValue

```ts
getTypedValue(sheetId: string, address: string): TypedValue;
```

Returns the calculated value with its type: `"number"`, `"text"`, `"boolean"`, `"blank"` or `"error"`. An unknown sheet or an address outside the grid returns `{ type: "error", value: "#REF!" }`. A formula that is waiting in manual mode returns `#PENDING!`.

```js
calculator.getTypedValue(sheet.id, "B1"); // { type: "number", value: 250 }
```

### evaluateFormula

```ts
evaluateFormula(sheetId: string, address: string, formula: string): unknown;
```

Evaluates a formula as if it were in `address`, without storing it. The result is the value itself, not a typed value. Errors come back as text such as `"#VALUE!"`, so compare with `getTypedValue` when the type matters. A formula that does not start with `=` returns `"#VALUE!"`. A formula with a syntax error returns `"#ERROR!"` and adds `#ERROR!` to `getOperationalErrors()`.

```js
calculator.evaluateFormula(sheet.id, "C1", "=B1*2"); // 500
calculator.evaluateFormula(sheet.id, "C1", "A1*2"); // "#VALUE!"
```

### getSpillEntries

```ts
getSpillEntries(sheetId: string): Array<{ address: string; anchor: string }>;
```

Lists each cell that a spilling array formula fills, with the anchor cell that produced it. Anchor cells themselves are not listed. A worksheet with no spills returns an empty array.

For `=SEQUENCE(2,2)` in `A1`, the result is `[{ address: "B1", anchor: "A1" }, { address: "A2", anchor: "A1" }, { address: "B2", anchor: "A1" }]`.

```js
calculator.getSpillEntries(sheet.id); // [] when nothing spills
```

### getOperationalErrors

```ts
getOperationalErrors(): string[];
```

Returns the operational error codes raised during the calculation, such as `"#CYCLE!"`. A workbook with a circular reference returns `["#CYCLE!"]`. A workbook with no operational errors returns `[]`.

```js
calculator.getOperationalErrors(); // []
```

### getDateContext

```ts
getDateContext(): CalculationDateContext;
```

Returns the calendar, date system, text date order and time zone in use, along with `calculatedAt`, `localDateTime`, `utcOffsetMinutes`, `todaySerial` and `nowSerial`. The values come from the workbook's `datePolicy`, which defaults to UTC, 1900 dates and MDY text dates.

```js
const { timeZone, dateSystem } = calculator.getDateContext(); // "UTC", "1900"
```

### Other members

| Method                                 | Returns                                                                                                                            |
| -------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------- |
| `getSpillRange(sheetId, address)`      | The range an anchor cell's spill occupies, such as `"A1:B2"`. `undefined` for other cells.                                         |
| `getSpillAnchor(sheetId, address)`     | The anchor cell of the spill this cell belongs to, or `undefined`.                                                                 |
| `getSpillDiagnostic(sheetId, address)` | Why the spill anchored at this cell is blocked, or `undefined`.                                                                    |
| `peekSpill(sheetId, address)`          | `{ range, diagnostic }` recorded at the cell. Either field may be missing.                                                         |
| `getResourceReport()`                  | `{ memoryBudgetBytes, workBudget, accountedBytes, workUnits }` for the run.                                                        |
| `getFinancialDiagnostics()`            | One entry per financial function call, with `functionName`, `convention`, `status`, `message`, `sheetId` and `address`.            |
| `getNamedFormulaDiagnostics()`         | One entry per formula that calls an unsupported function, with `sheetId`, `address`, `error` and `message`.                        |
| `getReferenceDiagnostics()`            | One entry per refused reference, such as an external workbook reference, with `sheetId`, `address`, `kind`, `error` and `message`. |
