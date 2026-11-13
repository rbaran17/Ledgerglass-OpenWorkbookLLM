/**
 * Type declarations for the Node entry (`@ledgerglass/workbook/node`): the workbook model,
 * calculation, files and the assistant protocol, without the interface.
 */
import type { CalculationReport, TypedValue, Workbook } from "./index";
import type {
  DelimitedExportOptions,
  DelimitedImportOptions,
} from "./sdk/files";

export type * from "./index";
export {
  PROPOSE_WORKBOOK_CHANGES_TOOL_NAME,
  RENDER_WORKBOOK_TOOL_NAME,
  WorkbookSpecError,
  applyCellPatchTransaction,
  buildAssistantSelectionContext,
  buildCanonicalWorkbookSnapshot,
  buildWorkbookFromSpec,
  createAgUiAdapter,
  createBlankWorkbook,
  createWorkbookToolResult,
  createActorEdits,
  createWorkbookToolSubscriber,
  describeWorkbookAsSpec,
  getWorkbookCalculation,
  normalizeWorkbookInput,
  proposeWorkbookChangesTool,
  readRenderWorkbookCall,
  recordWorkbookChange,
  renderWorkbookTool,
  resolveAssistantSelection,
  runAgUiAgent,
  shiftFormulaReferences,
  summarizeWorkbookForAgent,
  validateAssistantFormula,
  validateAssistantProposal,
  workbookLimits,
  workbookSpecLimits,
  workbookSpecSchema,
} from "./index";
export * from "./sdk/calculation";
export * from "./sdk/cells";
export * from "./sdk/files";
export * from "./sdk/rules";
export * from "./sdk/tables";
export * from "./sdk/worksheets";

export interface ReadWorkbookFileOptions {
  /** Parser options for .csv, .tsv and .txt files. */
  delimited?: DelimitedImportOptions;
  signal?: AbortSignal;
  timeoutMs?: number;
}

/**
 * Reads a workbook file by path. The extension picks the format: .json is native JSON (validated),
 * .csv/.tsv/.txt are delimited text, anything else goes through the workbook importer (XLSX).
 * Rejects with `FileProcessingError` for a file that is not accepted, with the file system's error
 * if the path cannot be read, and with an Error for native JSON that is not a workbook.
 */
export function readWorkbookFile(
  path: string,
  options?: ReadWorkbookFileOptions,
): Promise<{
  workbook: Workbook;
  /** What the format could not carry over. Empty for native JSON and delimited text. */
  warnings: string[];
  /** The importer's report; undefined for native JSON. */
  report: Record<string, unknown> | undefined;
}>;

export interface WriteWorkbookFileOptions {
  /** Sheet to write as delimited text. Defaults to the first sheet. */
  sheetId?: string;
  /** Options for .csv and .tsv. The delimiter defaults to a tab for .tsv. */
  delimited?: DelimitedExportOptions;
}

/**
 * Writes a workbook to a path. The extension picks the format: .xlsx, .json (native; keeps
 * everything) or .csv/.tsv (one sheet, formulas as values). Throws a TypeError for any other
 * extension.
 */
export function writeWorkbookFile(
  workbook: Workbook,
  path: string,
  options?: WriteWorkbookFileOptions,
): Promise<{
  path: string;
  /** Size of the file written. */
  bytes: number;
  /** What the format could not keep. */
  warnings: string[];
}>;

/**
 * Calculates the workbook and returns every filled cell's result as
 * `sheets[sheetName][address]`, with the calculation report.
 */
export function calculateWorkbookValues(workbook: Workbook): {
  report: CalculationReport;
  sheets: Record<string, Record<string, TypedValue>>;
};
