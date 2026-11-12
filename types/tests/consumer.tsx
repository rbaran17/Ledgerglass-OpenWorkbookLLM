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
  AgentWorkbook,
  RENDER_WORKBOOK_TOOL_NAME,
  WORKBOOK_COMPONENT_NAME,
  WORKBOOK_CATALOG_ID,
  WorkbookSpecError,
  buildWorkbookFromSpec,
  createWorkbookCatalog,
  createWorkbookToolResult,
  createWorkbookToolSubscriber,
  describeWorkbookAsSpec,
  readRenderWorkbookCall,
  renderWorkbookTool,
  summarizeWorkbookForAgent,
  workbookSpecLimits,
  type AgentWorkbookProps,
  type AssistantAdapter,
  type AssistantProposal,
  type Workbook,
  type WorkbookSheetSpec,
  type WorkbookSpec,
  type WorkbookToolMessage,
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

// Agent-rendered workbook: a host showing a description an agent sent.
const agentSpec: WorkbookSpec = {
  title: "Budget",
  sheets: [
    {
      name: "Plan",
      rows: [
        ["Item", "Cost"],
        ["Rent", 1200],
      ],
      cells: { B4: "=SUM(B2:B3)" },
    },
  ],
};
const agentBuilt = buildWorkbookFromSpec(agentSpec, { id: "call-1" });
const agentWorkbook: Workbook = agentBuilt.workbook;
agentBuilt.title satisfies string;
agentBuilt.readOnly satisfies boolean;
workbookSpecLimits.maxCells satisfies number;

const agentSummary = summarizeWorkbookForAgent(agentWorkbook, {
  maxErrors: 10,
});
agentSummary.formulaErrorCount satisfies number;
agentSummary.formulaErrors.map((entry) => entry.error satisfies string);

const agentDescription: WorkbookSpec = describeWorkbookAsSpec(agentWorkbook, {
  readOnly: true,
});
agentDescription.sheets satisfies WorkbookSheetSpec[];

// The AG-UI tool: a rejected call is answered with its problems, not a workbook.
renderWorkbookTool.name satisfies typeof RENDER_WORKBOOK_TOOL_NAME;
const agentCall = readRenderWorkbookCall('{"sheets":[{"name":"Plan"}]}', {
  toolCallId: "call-1",
});
if (agentCall.ok) agentCall.workbook.name satisfies string;
else agentCall.problems satisfies string[];
const agentRaw = readRenderWorkbookCall("{}");
// @ts-expect-error the call may have been rejected, so workbook is not there yet
agentRaw.workbook;

const agentMessage: WorkbookToolMessage = createWorkbookToolResult("call-1", {
  sheets: [{ name: "Plan" }],
});
agentMessage.role satisfies "tool";

const agentSubscriber = createWorkbookToolSubscriber({
  onWorkbook: (call) => {
    call.result.content satisfies string;
    call.spec.sheets satisfies WorkbookSheetSpec[];
  },
  onInvalid: (call) => call.problems satisfies string[],
});
agentSubscriber.onToolCallEndEvent({
  event: { toolCallId: "call-1" },
  toolCallName: RENDER_WORKBOOK_TOOL_NAME,
  toolCallArgs: "{}",
});

try {
  buildWorkbookFromSpec({ sheets: [] }, { id: "call-2" });
} catch (error) {
  if (error instanceof WorkbookSpecError) error.problems satisfies string[];
}

// The AG-UI host renders AgentWorkbook; edits come back as a description.
export const agentElement = (
  <AgentWorkbook
    spec={agentSpec}
    workbookId="call-1"
    height={420}
    readOnly={false}
    adapter={localDemoAdapter}
    actor="Agent"
    onChange={(next) => next.revision}
    onInvalid={(problems) => problems.length satisfies number}
    onSpecChange={(spec, changed) => {
      spec.sheets satisfies WorkbookSheetSpec[];
      changed.revision satisfies number;
    }}
  />
);
// @ts-expect-error the description is required
const agentPropsWithoutSpec: AgentWorkbookProps = { workbookId: "call-1" };
void agentPropsWithoutSpec;

// The A2UI host builds the catalog from the pieces the A2UI packages provide. A stand-in here.
class StandInCatalog {
  constructor(
    readonly catalogId: string,
    readonly protocolVersion: string,
    readonly components: unknown[],
    readonly functions: unknown[],
    readonly theme: unknown,
    readonly instructions: string,
  ) {}
  toJson() {
    return { catalogId: this.catalogId, instructions: this.instructions };
  }
}
const agentCatalog = createWorkbookCatalog(
  {
    Catalog: StandInCatalog,
    CommonSchemas: {},
    createComponentImplementation: () => ({}),
    z: {},
  },
  { height: 420, editorProps: { actor: "Agent" } },
);
agentCatalog.toJson().catalogId satisfies string;
WORKBOOK_COMPONENT_NAME satisfies "Workbook";
WORKBOOK_CATALOG_ID satisfies string;
// @ts-expect-error the A2UI pieces, including Catalog, are required
createWorkbookCatalog({
  CommonSchemas: {},
  createComponentImplementation: () => ({}),
  z: {},
});
