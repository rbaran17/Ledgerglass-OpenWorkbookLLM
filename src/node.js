/**
 * Node entry (`@ledgerglass/workbook/node`): the workbook model, calculation, files and the
 * assistant protocol, with no interface code. Use it in scripts, servers and tests.
 * Its TypeScript declarations are in types/node.d.ts.
 */
export {
  applyCellPatchTransaction,
  buildCanonicalWorkbookSnapshot,
  createBlankWorkbook,
  getWorkbookCalculation,
  normalizeWorkbookInput,
  shiftFormulaReferences,
} from "./app/03-workbook-model.js";
export { workbookLimits } from "./app/02-text-and-financial-functions.js";

export * from "./sdk/calculation.js";
export * from "./sdk/cells.js";
export * from "./sdk/files.js";
export * from "./sdk/rules.js";
export * from "./sdk/tables.js";
export * from "./sdk/worksheets.js";

/** Recording of edits: see docs/api/README.md, "Edits are recorded". */
export { createActorEdits, recordWorkbookChange } from "./sdk/recording.js";

/** Node conveniences: files by path, and all calculated values at once. */
export { readWorkbookFile, writeWorkbookFile } from "./node/files.js";
export { calculateWorkbookValues } from "./node/values.js";

/** Assistant protocol: the approved-selection context and the checks applied to proposals. */
export {
  buildAssistantSelectionContext,
  resolveAssistantSelection,
  validateAssistantFormula,
  validateAssistantProposal,
} from "./app/19b-assistant-protocol.js";

/** The AG-UI assistant adapter and its client, for agents driven from a script or a server. */
export {
  PROPOSE_WORKBOOK_CHANGES_TOOL_NAME,
  createAgUiAdapter,
  proposeWorkbookChangesTool,
  runAgUiAgent,
} from "./integrations/ag-ui-adapter.js";

/** The workbook description an agent sends (shared with the AG-UI tool and A2UI component). */
export {
  WorkbookSpecError,
  buildWorkbookFromSpec,
  describeWorkbookAsSpec,
  summarizeWorkbookForAgent,
  workbookSpecLimits,
  workbookSpecSchema,
} from "./agent-ui/workbook-spec.js";
export {
  RENDER_WORKBOOK_TOOL_NAME,
  createWorkbookToolResult,
  createWorkbookToolSubscriber,
  readRenderWorkbookCall,
  renderWorkbookTool,
} from "./agent-ui/ag-ui.js";
