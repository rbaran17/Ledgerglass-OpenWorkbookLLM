/**
 * Library entry: the editor component, the workbook model and the assistant protocol.
 * Import the stylesheet as well: `import "@ledgerglass/workbook/style.css"`.
 *
 * The names exported here were chosen for the rebuilt library; the original package's exact API
 * was not recovered. There are no TypeScript declarations.
 */
import "./styles/workbook.css";

/** The editor. Give its container a height; see docs/LIBRARY.md for the props. */
export { WorkbookSyncBoundary as AuditWorkbook } from "./app/56-main-worksheet-editor.js";

/** Assistant adapters: the built-in local demonstration and one that calls your server. */
export { Vj as localDemoAdapter } from "./app/55-external-source-refresh.js";
export { createHttpAdapter } from "./integrations/http-adapter.js";

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

/** Assistant protocol: the approved-selection context and the checks applied to proposals. */
export {
  buildAssistantSelectionContext,
  resolveAssistantSelection,
  validateAssistantFormula,
  validateAssistantProposal,
} from "./app/19b-assistant-protocol.js";
