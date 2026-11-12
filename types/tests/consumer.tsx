/**
 * Compiled, never run: `npm run test:types` fails if the declarations stop matching how the
 * library and the MCP server are meant to be used.
 */
import {
  AuditWorkbook,
  applyCellPatchTransaction,
  buildAssistantSelectionContext,
  createBlankWorkbook,
  createHttpAdapter,
  getWorkbookCalculation,
  localDemoAdapter,
  normalizeWorkbookInput,
  shiftFormulaReferences,
  validateAssistantProposal,
  workbookLimits,
  type AssistantAdapter,
  type AssistantProposal,
  type Workbook,
} from "@ledgerglass/workbook";
import { createWorkbookMcpServer } from "@ledgerglass/workbook/mcp";

const workbook: Workbook = createBlankWorkbook("Typed");
const sheet = workbook.sheets[0];
const edited = applyCellPatchTransaction(
  workbook,
  [{ sheetId: sheet.id, address: "A1", cell: { raw: "=6*7" } }],
  "Script",
  "Set a formula",
);
const { calculator, report } = getWorkbookCalculation(edited);
const value = calculator.getTypedValue(sheet.id, "A1");
if (value.type === "number") value.value satisfies number;
report.formulaCount satisfies number;
normalizeWorkbookInput(JSON.parse("{}")) satisfies Workbook;
shiftFormulaReferences("=A1", 1, 0) satisfies string;
workbookLimits.maxLLMCells satisfies number;

const context = buildAssistantSelectionContext(edited, {
  sheetId: sheet.id,
  range: "A1:B2",
});
const checked: AssistantProposal = validateAssistantProposal({}, context);

const adapter: AssistantAdapter = createHttpAdapter({
  endpoint: "/api/assistant",
});
const custom: AssistantAdapter = {
  name: "Custom",
  async propose(prompt, selection) {
    return {
      id: "p",
      workbookId: selection.workbookId,
      revision: selection.revision,
      explanation: prompt,
      citations: [],
      patches: [],
    };
  },
};

export const element = (
  <AuditWorkbook
    initialWorkbook={edited}
    adapter={custom ?? adapter ?? localDemoAdapter}
    actor="Reviewer"
    readOnly={false}
    onChange={(next) => next.revision}
  />
);

const queue: AssistantProposal[] = [];
export const server = createWorkbookMcpServer({
  getWorkbook: () => edited,
  getSelection: () => ({ sheetId: sheet.id, range: "A1:B2" }),
  queueProposal: (proposal) => queue.push(proposal),
});
void checked;

// @ts-expect-error a patch needs a cell
applyCellPatchTransaction(workbook, [{ sheetId: sheet.id, address: "A1" }]);
// @ts-expect-error an adapter needs an endpoint
createHttpAdapter({});
// @ts-expect-error the host must supply a proposal queue
createWorkbookMcpServer({
  getWorkbook: () => edited,
  getSelection: () => ({ sheetId: "", range: "" }),
});
