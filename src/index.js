/**
 * Library entry: the editor component, the workbook model and the assistant protocol.
 * Import the stylesheet as well: `import "@ledgerglass/workbook/style.css"`.
 *
 * The names exported here were chosen for the rebuilt library; the original package's exact API
 * was not recovered. Their TypeScript declarations are in types/index.d.ts.
 */
import "./styles/workbook.css";

/** The editor. Give its container a height; see docs/LIBRARY.md for the props. */
export { WorkbookSyncBoundary as AuditWorkbook } from "./app/56-main-worksheet-editor.js";

/** Assistant adapters: the built-in local demonstration and one that calls your server. */
export { Vj as localDemoAdapter } from "./app/55-external-source-refresh.js";
export { createHttpAdapter } from "./integrations/http-adapter.js";

/**
 * External data sources: the adapter for the editor's `externalSourceAdapter` prop, which talks
 * to a source gateway on your server (see docs/SOURCES.md), and the error it reports.
 */
export { createHttpSourceAdapter } from "./integrations/http-source-adapter.js";
export { ExternalSourceError } from "./app/03-workbook-model.js";

/** Workbook model: create, normalise, calculate and edit workbooks without the interface. */
export {
  applyCellPatchTransaction,
  buildCanonicalWorkbookSnapshot,
  createBlankWorkbook,
  getWorkbookCalculation,
  normalizeWorkbookInput,
  shiftFormulaReferences,
} from "./app/03-workbook-model.js";
export { workbookLimits } from "./app/02-text-and-financial-functions.js";

/**
 * The wider model API, by area. Each area has a reference page in docs/api/ and its declarations
 * in types/sdk/. These functions also load in Node through `@ledgerglass/workbook/node`.
 */
export * from "./sdk/calculation.js";
export * from "./sdk/cells.js";
export * from "./sdk/files.js";
export * from "./sdk/rules.js";
export * from "./sdk/tables.js";
export * from "./sdk/worksheets.js";

/** Assistant protocol: the approved-selection context and the checks applied to proposals. */
export {
  buildAssistantSelectionContext,
  resolveAssistantSelection,
  validateAssistantFormula,
  validateAssistantProposal,
} from "./app/19b-assistant-protocol.js";

/**
 * Agent-rendered workbook: lets an agent show the editor through AG-UI (a frontend tool) or A2UI
 * (a catalog component). See docs/AGENT-UI.md.
 */
export { AgentWorkbook } from "./agent-ui/AgentWorkbook.js";
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
export {
  WORKBOOK_CATALOG_ID,
  WORKBOOK_COMPONENT_NAME,
  createWorkbookCatalog,
  workbookCatalogInstructions,
} from "./agent-ui/a2ui.js";
