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

/** The size limits the model enforces: rows, columns, cells, file sizes and assistant selection sizes. */
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

/** Returns a new workbook with one empty sheet. */
export function createBlankWorkbook(name?: string): Workbook;
/** Validates and completes workbook data, for example parsed native JSON. Throws if it is invalid. */
export function normalizeWorkbookInput(input: unknown): Workbook;
/** Calculates the workbook (or reuses the cached calculation) and returns the calculator and its report. */
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
/** Returns a normalised copy of the workbook suitable for comparing or hashing. */
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
/** Builds what an assistant may see for an approved selection: the cells, their values and identifiers. Throws if the selection is not allowed. */
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

/** Returns an assistant adapter that asks your own server for a proposal. */
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

/*
 * Agent-rendered workbook: an agent shows the editor through AG-UI (a frontend tool) or A2UI (a
 * catalog component). See docs/AGENT-UI.md.
 */

/** A cell in an agent's description. Text starting with "=" is a formula. */
export type WorkbookSpecCell = string | number | boolean | null;

export interface WorkbookSheetSpec {
  /** Tab name: unique, at most 31 characters, without : \ / ? * [ ]. */
  name: string;
  /** Cell values row by row, starting at `startCell`. The first row is usually the header. */
  rows?: WorkbookSpecCell[][];
  /** Top-left cell of `rows`. Defaults to "A1". */
  startCell?: string;
  /** Single cells by address, for example `{ "D7": "=SUM(D2:D6)" }`. Applied after `rows`. */
  cells?: Record<string, WorkbookSpecCell>;
}

/** The description an agent sends to show a workbook. */
export interface WorkbookSpec {
  /** Short name, shown above the grid. */
  title?: string;
  /** The worksheets, in tab order. */
  sheets: WorkbookSheetSpec[];
  /** True to show the workbook without allowing edits. Defaults to false. */
  readOnly?: boolean;
}

/** Limits for agent-supplied workbooks. */
export const workbookSpecLimits: Readonly<{
  maxSheets: number;
  /** Filled cells across all sheets. */
  maxCells: number;
  /** Longest text in one cell. */
  maxTextLength: number;
  maxTitleLength: number;
  maxSheetNameLength: number;
}>;

/** JSON Schema of `WorkbookSpec`, used as the AG-UI tool parameters and in the A2UI catalog. */
export const workbookSpecSchema: Readonly<Record<string, unknown>>;

export interface WorkbookSpecErrorInstance extends Error {
  /** Every reason the description was rejected, in plain words. */
  problems: string[];
}

/** Thrown when a description cannot be shown. `problems` lists every reason. */
export const WorkbookSpecError: new (
  problems: string[],
) => WorkbookSpecErrorInstance;

/**
 * Builds a workbook from an agent's description. The input is untrusted, so it is typed `unknown`.
 * Throws `WorkbookSpecError` listing every problem. Pass a stable `id` so that showing the same
 * description again keeps the same document.
 */
export function buildWorkbookFromSpec(
  spec: unknown,
  options?: { id?: string },
): { workbook: Workbook; title: string; readOnly: boolean };

/** A compact account of a workbook for the agent. */
export interface WorkbookAgentSummary {
  workbookId: string;
  title: string;
  sheets: Array<{ name: string; filledCells: number; formulaCount: number }>;
  /** Every formula that produced an error, counted. */
  formulaErrorCount: number;
  /** The first `maxErrors` of them (25 by default). */
  formulaErrors: Array<{
    sheet: string;
    address: string;
    formula: string;
    /** The error text, such as "#DIV/0!". */
    error: string;
  }>;
}

/** What was shown and which formulas produced errors, so the agent can correct them. */
export function summarizeWorkbookForAgent(
  workbook: Workbook,
  options?: { maxErrors?: number },
): WorkbookAgentSummary;

/** Turns a workbook back into a description, for example to report the user's edits. */
export function describeWorkbookAsSpec(
  workbook: Workbook,
  options?: { readOnly?: boolean },
): WorkbookSpec;

/** Props of `AgentWorkbook`: the editor's props, except `initialWorkbook`, plus the description. */
export interface AgentWorkbookProps extends Pick<
  AuditWorkbookProps,
  | "onChange"
  | "onWorkbookLoaded"
  | "onSelectionChange"
  | "adapter"
  | "actor"
  | "readOnly"
  | "persistenceKey"
  | "workbookStore"
> {
  spec: WorkbookSpec;
  /** Stable identifier, such as the tool call or surface identifier. */
  workbookId?: string;
  /** Height in pixels, or a CSS length. Defaults to 480. */
  height?: number | string;
  /** Called after the user edits, with the edited workbook as a description and as a workbook. */
  onSpecChange?(spec: WorkbookSpec, workbook: Workbook): void;
  /** Called with the reasons when the description is rejected. */
  onInvalid?(problems: string[]): void;
  [extra: string]: unknown;
}

/** Shows a workbook from an agent's description, or the reasons it cannot be shown. */
export const AgentWorkbook: ComponentType<AgentWorkbookProps>;

/** The name of the AG-UI frontend tool. */
export const RENDER_WORKBOOK_TOOL_NAME: "render_workbook";

export interface WorkbookToolDefinition {
  name: "render_workbook";
  description: string;
  parameters: Readonly<Record<string, unknown>>;
}

/** The tool definition for `RunAgentInput.tools`. */
export const renderWorkbookTool: Readonly<WorkbookToolDefinition>;

export type RenderWorkbookCall =
  | {
      ok: true;
      spec: WorkbookSpec;
      workbook: Workbook;
      title: string;
      readOnly: boolean;
    }
  | { ok: false; problems: string[] };

/**
 * Reads the arguments of a `render_workbook` call: JSON text, or an object already parsed.
 * Check `ok` before using `workbook`.
 */
export function readRenderWorkbookCall(
  toolCallArgs: unknown,
  options?: { toolCallId?: string },
): RenderWorkbookCall;

/** The tool message that answers a `render_workbook` call. */
export interface WorkbookToolMessage {
  id: string;
  role: "tool";
  toolCallId: string;
  /** JSON text: `{ shown: true, ...summary }`, or `{ shown: false, problems }`. */
  content: string;
  /** Set when the description was rejected. */
  error?: string;
}

/**
 * The tool message for a call. Add it to the `messages` of the next run. A rejected description is
 * answered with `error` set, so the agent can correct it.
 */
export function createWorkbookToolResult(
  toolCallId: string,
  toolCallArgs: unknown,
  options?: { messageId?: string },
): WorkbookToolMessage;

/** An AG-UI subscriber that reports each finished `render_workbook` call. */
export interface WorkbookToolSubscriber {
  onToolCallEndEvent(params: {
    event: { toolCallId: string };
    toolCallName: string;
    toolCallArgs: unknown;
  }): void;
}

export interface WorkbookToolSubscriberHandlers {
  onWorkbook?(call: {
    toolCallId: string;
    ok: true;
    spec: WorkbookSpec;
    workbook: Workbook;
    title: string;
    readOnly: boolean;
    /** The tool message to add to the next run. */
    result: WorkbookToolMessage;
  }): void;
  onInvalid?(call: {
    toolCallId: string;
    problems: string[];
    /** The tool message to add to the next run. */
    result: WorkbookToolMessage;
  }): void;
}

/** Use with `agent.subscribe(...)` or `agent.runAgent(parameters, subscriber)`. */
export function createWorkbookToolSubscriber(
  handlers?: WorkbookToolSubscriberHandlers,
): WorkbookToolSubscriber;

/** Identifies the catalog to agents and renderers. It is a name; nothing is fetched from it. */
export const WORKBOOK_CATALOG_ID: string;

/** The name of the catalog component. */
export const WORKBOOK_COMPONENT_NAME: "Workbook";

/** Instructions for the agent about the Workbook component. */
export const workbookCatalogInstructions: string;

export interface WorkbookCatalogOptions {
  /** Defaults to `WORKBOOK_CATALOG_ID`. */
  catalogId?: string;
  /** Defaults to "1.0". */
  protocolVersion?: string;
  /** Default height in pixels. Defaults to 480. */
  height?: number;
  /** Props for every editor the catalog renders (adapter, actor, ...). */
  editorProps?: Record<string, unknown>;
}

/**
 * Builds the A2UI catalog that holds the Workbook component. The host passes in the A2UI pieces it
 * already has, so this library does not depend on the A2UI packages. Returns the host's own
 * catalog type.
 */
export function createWorkbookCatalog<TCatalog>(
  a2ui: {
    /** `Catalog` from `@a2ui/web_core/v1_0`. */
    Catalog: new (...args: any[]) => TCatalog;
    /** `CommonSchemas` from `@a2ui/web_core/v1_0`. */
    CommonSchemas: Record<string, any>;
    /** `createComponentImplementation` from `@a2ui/react`. */
    createComponentImplementation: (...args: any[]) => unknown;
    /** The `zod` the A2UI packages use (zod 3 for A2UI 0.15). */
    z: any;
  },
  options?: WorkbookCatalogOptions,
): TCatalog;

/** The wider model API, by area. Reference pages are in docs/api/. */
export * from "./sdk/calculation";
export * from "./sdk/cells";
export * from "./sdk/files";
export * from "./sdk/rules";
export * from "./sdk/tables";
export * from "./sdk/worksheets";
