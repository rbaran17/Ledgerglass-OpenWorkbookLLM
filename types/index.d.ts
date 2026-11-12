/**
 * Type declarations for the Ledgerglass library (`@ledgerglass/workbook`).
 *
 * Written by hand for the rebuilt library and checked against real use in `types/tests`. They
 * describe the public surface; fields the declarations do not name are still present at run time,
 * which is why several object types allow extra properties.
 */
import type { ComponentType } from "react";

/** What a cell stores: a number, text, a boolean, a formula beginning with "=", or nothing. */
export type RawCellContent = string | number | boolean | null;

export interface Cell {
  raw: RawCellContent;
  /** Excel number format code, such as "0.00". */
  numFmt?: string;
  /** Named format: "number", "currency", "percent" and others. */
  format?: string;
  style?: Record<string, unknown>;
  note?: string;
  [extra: string]: unknown;
}

export interface Sheet {
  /** Stable identifier; it does not change when the sheet is renamed. */
  id: string;
  name: string;
  /** Cells by address, for example `cells.B2`. Empty cells are absent. */
  cells: Record<string, Cell>;
  rowCount: number;
  colCount: number;
  frozenRows: number;
  frozenCols?: number;
  hidden?: boolean;
  [extra: string]: unknown;
}

export interface WorkbookEvent {
  id: string;
  timestamp: string;
  actor: string;
  action: string;
  changes: Array<{
    sheetId: string;
    address: string;
    before?: Cell;
    after?: Cell;
  }>;
  [extra: string]: unknown;
}

export interface Workbook {
  schemaVersion: 1;
  id: string;
  name: string;
  /** Increases by one with every committed change. */
  revision: number;
  sheets: Sheet[];
  events: WorkbookEvent[];
  metadata: { createdAt: string; modifiedAt: string; [extra: string]: unknown };
  [extra: string]: unknown;
}

/** A calculated value with its type. Errors carry their text, such as "#DIV/0!". */
export type TypedValue =
  | { type: "number"; value: number }
  | { type: "text"; value: string }
  | { type: "boolean"; value: boolean }
  | { type: "error"; value: string }
  | { type: "blank"; value: null };

export interface Calculator {
  getTypedValue(sheetId: string, address: string): TypedValue;
  /** Calculates a formula as if it were in `address`, without storing it. */
  evaluateFormula(sheetId: string, address: string, formula: string): unknown;
  getSpillEntries(sheetId: string): Array<{ address: string; anchor: string }>;
  getOperationalErrors(): string[];
  getDateContext(): Record<string, unknown>;
}

export interface CalculationReport {
  workbookId: string;
  revision: number;
  formulaCount: number;
  evaluated: number;
  cycles: unknown[];
  execution: { mode: string; state: string; message: string };
  [extra: string]: unknown;
}

export interface CellPatch {
  sheetId: string;
  address: string;
  cell: Cell;
}

export const workbookLimits: Readonly<{
  maxRows: number;
  maxCols: number;
  maxCells: number;
  maxFileBytes: number;
  maxNativeBytes: number;
  /** Largest selection an assistant may be given. */
  maxLLMCells: number;
  /** Most cell changes in one transaction or proposal. */
  maxPatches: number;
}>;

export function createBlankWorkbook(name?: string): Workbook;
/** Validates and completes workbook data, for example parsed native JSON. Throws if it is invalid. */
export function normalizeWorkbookInput(input: unknown): Workbook;
export function getWorkbookCalculation(workbook: Workbook): {
  calculator: Calculator;
  report: CalculationReport;
};
/** Returns a new workbook with the changes applied and recorded; the original is not modified. */
export function applyCellPatchTransaction(
  workbook: Workbook,
  patches: CellPatch[],
  actor?: string,
  action?: string,
  options?: Record<string, unknown>,
): Workbook;
export function buildCanonicalWorkbookSnapshot(workbook: Workbook): Workbook;
/** Moves the relative references in a formula. A reference pushed off the grid becomes "#REF!". */
export function shiftFormulaReferences(
  formula: string,
  rowOffset: number,
  columnOffset: number,
): string;

/** The approved selection an assistant may see. */
export interface AssistantSelection {
  sheetId: string;
  /** One rectangle, such as "A1:D20". */
  range: string;
}

export interface AssistantContext {
  workbookId: string;
  revision: number;
  sheetId: string;
  sheetName: string;
  range: string;
  cells: Array<{ address: string; raw: RawCellContent; value: unknown }>;
  [extra: string]: unknown;
}

export interface AssistantProposal {
  id: string;
  workbookId: string;
  revision: number;
  explanation: string;
  /** Addresses inside the approved selection that support the proposal. */
  citations: string[];
  patches: CellPatch[];
  grounding?: unknown;
}

/** Cell addresses of the selection, row by row. Throws if the selection is not allowed. */
export function resolveAssistantSelection(
  workbook: Workbook,
  selection: AssistantSelection,
): string[];
export function buildAssistantSelectionContext(
  workbook: Workbook,
  selection: AssistantSelection,
): AssistantContext;
/** Throws unless the formula uses allowed functions and only cells in `approvedAddresses`. */
export function validateAssistantFormula(
  formula: string,
  approvedAddresses: Set<string>,
): void;
/** Returns the proposal in checked form. Throws with a plain-language reason if it is not acceptable. */
export function validateAssistantProposal(
  proposal: unknown,
  context: AssistantContext,
): AssistantProposal;

export interface AssistantAdapter {
  name: string;
  /** True when the adapter never sends data anywhere. */
  localOnly?: boolean;
  chat?: { history?: boolean; attachments?: boolean };
  propose(
    prompt: string,
    context: AssistantContext,
    signal?: AbortSignal,
    request?: unknown,
  ): Promise<AssistantProposal>;
}

/** The built-in demonstration adapter. It calls no model and no network. */
export const localDemoAdapter: AssistantAdapter;

export function createHttpAdapter(options: {
  /** Your server's URL. It receives `{ prompt, context, request }` and returns a proposal. */
  endpoint: string;
  name?: string;
  chat?: boolean;
  headers?: Record<string, string>;
  fetch?: typeof fetch;
}): AssistantAdapter;

export interface AuditWorkbookProps {
  /** The workbook to open. Change its `id` to switch documents. */
  initialWorkbook?: Workbook;
  /** Called after each in-memory change. Not a durable-save acknowledgement. */
  onChange?(workbook: Workbook): void;
  onWorkbookLoaded?(workbook: Workbook): void;
  onSelectionChange?(selection: unknown): void;
  adapter?: AssistantAdapter;
  /** Name recorded against edits. Defaults to "Preparer". */
  actor?: string;
  readOnly?: boolean;
  /** Browser storage key for saving locally. Without it nothing is stored. */
  persistenceKey?: string;
  workbookStore?: unknown;
  [extra: string]: unknown;
}

/** The editor. It fills its container, so give the container a height. */
export const AuditWorkbook: ComponentType<AuditWorkbookProps>;
