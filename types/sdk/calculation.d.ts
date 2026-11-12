/**
 * Type declarations for the calculation area of `@ledgerglass/workbook`. Each name matches an export of `src/sdk/calculation.js`.
 */
import type { CalculationReport, TypedValue, Workbook } from "../index";

/** Calculation settings. Unknown keys are rejected. `memoryBudgetMiB` and `workBudget` are optional. */
export interface CalculationSettings {
  /** "automatic" recalculates on every committed change; "manual" waits for an explicit refresh. */
  mode: "automatic" | "manual";
  /** When true, circular references are iterated instead of returning #CYCLE!. */
  iterative: boolean;
  /** Integer from 1 to 1000. */
  maxIterations: number;
  /** Convergence threshold from 0 to 1. */
  maxChange: number;
  /** Integer from 1 to 10000 milliseconds. */
  timeBudgetMs: number;
  /** Integer from 1 to 128 (mebibytes of working data). */
  memoryBudgetMiB?: number;
  /** Integer from 1000 to 10000000 work units. */
  workBudget?: number;
}

/** Limits derived from calculation settings, in bytes and work units. */
export interface CalculationBudget {
  memoryBudgetBytes: number;
  workBudget: number;
}

/** Resource use recorded by one calculation run. */
export interface CalculationResourceReport {
  memoryBudgetBytes: number;
  workBudget: number;
  accountedBytes: number;
  workUnits: number;
  [extra: string]: unknown;
}

/** Zero-based, inclusive row interval. */
export interface RowInterval {
  start: number;
  end: number;
}

/** Row intervals that the calculation treats as filtered out, keyed by worksheet ID. */
export type FilteredRowsBySheet = Record<string, RowInterval[]>;

/** The calculation context: calendar, time zone and the serial dates used by NOW and TODAY. */
export interface CalculationDateContext {
  calendar: string;
  dateSystem: string;
  textDateOrder: string;
  timeZone: string;
  calculatedAt: string;
  localDateTime: string;
  utcOffsetMinutes: number;
  todaySerial: number;
  nowSerial: number;
  [extra: string]: unknown;
}

/** An array formula that spills onto cells other than its anchor. */
export interface SpillEntry {
  address: string;
  anchor: string;
}

/** Why a spilling formula could not place its result. */
export interface SpillDiagnostic {
  sheetId: string;
  anchor: string;
  reason: string;
  obstruction?: string;
  message: string;
  [extra: string]: unknown;
}

/** A financial function result and the sign, timing and day-count convention it used. */
export interface FinancialDiagnostic {
  functionName: string;
  convention: string;
  status: string;
  message: string;
  sheetId: string;
  address: string;
  [extra: string]: unknown;
}

/** A formula that calls a function the calculator does not run. */
export interface NamedFormulaDiagnostic {
  sheetId: string;
  address: string;
  error: string;
  message: string;
  [extra: string]: unknown;
}

/** A reference the calculator refuses, such as an external workbook reference. */
export interface ReferenceDiagnostic {
  sheetId: string;
  address: string;
  kind: string;
  error: string;
  message: string;
  [extra: string]: unknown;
}

/**
 * The calculator returned by `getWorkbookCalculation(workbook).calculator` and by the
 * calculation functions in this area. It is a function as well as an object: calling it with a
 * sheet ID and an address returns the cell's value, the same as `getTypedValue(...).value`.
 */
export interface CalculatorMethods {
  (sheetId: string, address: string): unknown;
  /** The calculated value of a cell with its type. Addresses are case-insensitive and may contain `$`. Unknown sheets and addresses past the grid return `#REF!`. */
  getTypedValue(sheetId: string, address: string): TypedValue;
  /** Evaluates a formula as if it stood in `address`, without storing it. Returns a scalar; errors come back as text such as `"#VALUE!"`. Formulas that do not start with `=` return `"#VALUE!"`. */
  evaluateFormula(sheetId: string, address: string, formula: string): unknown;
  /** The rectangle that the spill anchored at this cell occupies, such as `"A1:B2"`. Undefined for any cell that is not a spill anchor. */
  getSpillRange(sheetId: string, address: string): string | undefined;
  /** Why the spill anchored at this cell is blocked, or undefined when it is not blocked. */
  getSpillDiagnostic(
    sheetId: string,
    address: string,
  ): SpillDiagnostic | undefined;
  /** The anchor cell of the spill this cell belongs to, or undefined. */
  getSpillAnchor(sheetId: string, address: string): string | undefined;
  /** Every spilled cell on a worksheet, with the anchor that produced it. Anchors themselves are not listed. */
  getSpillEntries(sheetId: string): SpillEntry[];
  /** The resource use of the last run: budgets, accounted bytes and work units. */
  getResourceReport(): CalculationResourceReport;
  /** Operational error codes raised during the calculation, such as `#CYCLE!`. Calling `evaluateFormula` with a formula that fails to parse adds `#ERROR!` to this list. */
  getOperationalErrors(): string[];
  /** The calendar, time zone and serial dates in use for NOW and TODAY. */
  getDateContext(): CalculationDateContext;
  /** One entry per financial function call, with the convention it used. */
  getFinancialDiagnostics(): FinancialDiagnostic[];
  /** The spill range and spill diagnostic recorded at a cell. */
  peekSpill(
    sheetId: string,
    address: string,
  ): { range?: string; diagnostic?: SpillDiagnostic };
  /** One entry per formula that calls an unsupported function. */
  getNamedFormulaDiagnostics(): NamedFormulaDiagnostic[];
  /** One entry per reference the calculator refused. */
  getReferenceDiagnostics(): ReferenceDiagnostic[];
}

/** A calculator together with the report that describes its run. */
export interface CalculationResult {
  calculator: CalculatorMethods;
  report: CalculationReport;
}

/** Row filters for a calculation. Keys are worksheet IDs. */
export interface CalculationOptions {
  filteredRowsBySheet?: FilteredRowsBySheet;
}

/** The fixed defaults: automatic mode, no iteration, 100 iterations, 0.001 convergence and a 2500 ms budget. */
export const defaultCalculationSettings: Readonly<CalculationSettings>;

/** Validates calculation settings and returns a normalized copy. Throws on any unknown key or out-of-range value. */
export function normalizeCalculationSettings(
  settings: unknown,
): CalculationSettings;

/** Returns a new workbook object with the normalized settings in `calculation`. Sheets are shared with the argument, and the argument is not modified. Throws when the settings are invalid. */
export function withCalculationSettings(
  workbook: Workbook,
  settings: CalculationSettings,
): Workbook;

/** Calculates a workbook with row filters applied and returns the calculator. Only `filteredRowsBySheet` is accepted; other keys throw. Filtered rows change SUBTOTAL results but not SUM. */
export function calculateWorkbookWithSettings(
  workbook: Workbook,
  options?: CalculationOptions,
): CalculatorMethods;

/** Recalculates every formula now, even if a current result exists, and stores the new result. Use it to refresh volatile functions such as RAND and manual-mode results. Returns the new calculator and report. */
export function refreshCalculationReport(workbook: Workbook): CalculationResult;

/** Returns the stored calculator and report for this exact workbook object, or undefined when none is current. It never calculates. */
export function getCurrentCalculatorRecord(
  workbook: Workbook,
): CalculationResult | undefined;

/** Returns the calculation report as a JSON string with the boundaries of the evidence. Throws when `report` is not the current report for `workbook`. */
export function buildCalculationEvidence(
  workbook: Workbook,
  report: CalculationReport,
): string;

/** Every built-in worksheet function name the calculator evaluates, in uppercase. */
export const builtinFunctionNames: ReadonlyArray<string>;

/** The statistical function names, such as AVERAGE and MEDIAN. Each is also in `builtinFunctionNames`. */
export const statisticalFunctionNames: ReadonlyArray<string>;

/** The database function names, such as DSUM and DGET. */
export const databaseFunctionNames: ReadonlyArray<string>;

/** Function names that return a matrix (array), such as SEQUENCE and FILTER. */
export const matrixResultFunctionNames: ReadonlyArray<string>;

/** The DAX function names the data model accepts, mapped to `[minimum, maximum]` argument counts. */
export const daxFunctionNames: Readonly<
  Record<string, readonly [number, number]>
>;

/** Function names in the text-cleanup group, such as TRIM and CLEAN. */
export const textCleanupFunctionNames: ReadonlyArray<string>;

/** The date function names, such as DATE and EDATE. */
export const dateFunctionNames: ReadonlyArray<string>;

/** The financial function names the calculator supports, such as PMT and NPV. */
export const supportedFinancialFunctionNames: ReadonlyArray<string>;

/** The text function names the calculator supports, such as LEFT and TEXTJOIN. */
export const supportedTextFunctionNames: ReadonlyArray<string>;

/** Converts calculation settings into byte and work limits. Missing settings use 32 MiB and 1,000,000 work units. */
export function resolveCalculationBudget(calculation?: {
  memoryBudgetMiB?: number;
  workBudget?: number;
}): CalculationBudget;

/** Estimates the working data a workbook needs, in bytes. Stops counting once the total passes `memoryBudgetBytes`, so an estimate above the budget is not the exact total. */
export function estimateWorkbookMemoryBytes(
  workbook: Workbook,
  memoryBudgetBytes: number,
): number;

/** Fixed limits for worksheet hidden intervals and frozen panes. */
export const worksheetCapacityLimits: Readonly<{
  maxIntervalsPerAxis: number;
  maxFrozenVisibleRows: number;
  maxFrozenVisibleCols: number;
  maxXlsxHiddenRows: number;
}>;

/** Fixed limits for what-if analyses: inputs, outputs, scenarios, axis values, trials and work. */
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

/** A cell that a what-if analysis changes or reads. */
export interface WhatIfCellReference {
  sheetId: string;
  /** A1-style address; it is stored normalized, for example `B2`. */
  address: string;
}

/** A named set of input values, one per changing input. */
export interface WhatIfScenario {
  name: string;
  values: number[];
}

/** A validated what-if definition with its changing inputs and formula results. */
export type WhatIfDefinition =
  | {
      name: string;
      kind: "scenarios";
      inputs: WhatIfCellReference[];
      outputs: WhatIfCellReference[];
      scenarios: WhatIfScenario[];
    }
  | {
      name: string;
      kind: "one-variable";
      inputs: WhatIfCellReference[];
      outputs: WhatIfCellReference[];
      orientation: "row" | "column";
      values: number[];
    }
  | {
      name: string;
      kind: "two-variable";
      inputs: WhatIfCellReference[];
      outputs: WhatIfCellReference[];
      rowValues: number[];
      columnValues: number[];
    };

/** One case of a what-if analysis: a display name and the input values to calculate. */
export interface WhatIfCase {
  name: string;
  inputs: number[];
}

/** Checks a what-if request and returns a normalized definition. Throws with a message that says what is wrong. */
export function validateWhatIfDefinition(request: unknown): WhatIfDefinition;

/** Expands a validated definition into its cases. Pass the result of `validateWhatIfDefinition`; this function does not check its input. */
export function expandWhatIfScenarios(request: WhatIfDefinition): WhatIfCase[];
