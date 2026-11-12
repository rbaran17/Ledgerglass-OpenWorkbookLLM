/**
 * Example: the Ledgerglass MCP server and an MCP client in the same Node process.
 *
 * The host builds a small workbook in code and approves one selection (A1:C4).
 * It creates the MCP server with createWorkbookMcpServer and connects a client to
 * it through the SDK's in-memory transport pair. The client reads the selection,
 * evaluates and audits it, then queues a proposal. The proposal lands in the
 * host's review queue; the workbook itself is never changed.
 *
 * Run from the repository root:
 *   node examples/mcp-client/01-in-process.mjs
 */
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import {
  applyCellPatchTransaction,
  createBlankWorkbook,
} from "@ledgerglass/workbook/node";
import { createWorkbookMcpServer } from "@ledgerglass/workbook/mcp";

// 1. Build the host's workbook. A blank workbook is changed through a recorded
//    transaction, which returns a new workbook; the original is not modified.
const blank = createBlankWorkbook("Ledger");
const sheetId = blank.sheets[0].id;
const cells = {
  A1: "Item",
  B1: "Amount",
  A2: "Rent",
  B2: 1200,
  A3: "Power",
  B3: "85.50", // a number typed as text: the audit should flag it
  A4: "Total",
  B4: "=B2+B3",
  C4: "=1/0", // a formula that fails: the audit should flag it
};
const patches = Object.entries(cells).map(([address, raw]) => ({
  sheetId,
  address,
  cell: { raw },
}));
const workbook = applyCellPatchTransaction(
  blank,
  patches,
  "Example",
  "Write the ledger",
);
const before = JSON.stringify(workbook);

// 2. The host decides what the model may see (the selection) and where proposals
//    go (its own review queue). The server cannot widen the selection.
const selection = { sheetId, range: "A1:C4" };
const reviewQueue = [];
const server = createWorkbookMcpServer({
  getWorkbook: () => workbook,
  getSelection: () => selection,
  queueProposal: (proposal) => {
    reviewQueue.push(proposal);
  },
});

// 3. Connect a client over the in-memory transport pair.
const client = new Client({ name: "example-host", version: "1.0.0" });
const [clientSide, serverSide] = InMemoryTransport.createLinkedPair();
await server.connect(serverSide);
await client.connect(clientSide);

// Calls a tool and returns its structured result; a failed call throws.
async function call(name, args = {}) {
  const result = await client.callTool({ name, arguments: args });
  if (result.isError) throw new Error(`${name}: ${result.content[0].text}`);
  return result.structuredContent;
}

const { tools } = await client.listTools();
console.log(`tools (${tools.length}):`);
for (const name of tools.map((tool) => tool.name).sort())
  console.log(`  ${name}`);

const capabilities = await call("workbook_get_capabilities");
console.log(`max selection cells: ${capabilities.limits.maxSelectionCells}`);
console.log(`writes: ${capabilities.authority.writes}`);

const context = await call("workbook_get_context");
console.log(
  `context: sheet ${context.sheetName}, range ${context.range}, ` +
    `${context.cells.length} cells, revision ${context.revision}`,
);

const read = await call("workbook_read_range", { range: "A2:B2" });
console.log(`read A2:B2: ${read.cells.map((cell) => cell.raw).join(" | ")}`);

const inspected = await call("workbook_inspect_formula", {
  formula: "=SUM(B2:B3)",
});
console.log(
  `inspect =SUM(B2:B3): functions ${inspected.functions.join(", ")}, ` +
    `references ${inspected.references.join(", ")}, ` +
    `proposable ${inspected.proposable}`,
);

const evaluated = await call("workbook_evaluate_formula", {
  formula: "=B2*2",
  address: "B4",
});
console.log(`evaluate =B2*2 as B4: ${evaluated.result}`);

const audit = await call("workbook_run_audit_checks");
console.log(
  `audit: ${audit.cellsChecked} cells checked, ${audit.findings.length} findings`,
);
for (const finding of audit.findings)
  console.log(`  ${finding.address} ${finding.check} (${finding.severity})`);

// 4. Queue a proposal. It must cite the revision the model read from the context.
const proposal = {
  id: "example-1",
  workbookId: context.workbookId,
  revision: context.revision,
  explanation: "Store the power amount as a number so the total can use it.",
  citations: ["B3"],
  patches: [{ sheetId, address: "B3", cell: { raw: 85.5 } }],
};
const queued = await call("workbook_queue_proposal", { proposal });
console.log(`queued: ${queued.queued}, applied: ${queued.applied}`);
console.log(`host review queue: ${reviewQueue.length} proposal`);
console.log(
  `queued patches: ${reviewQueue[0].patches
    .map((patch) => `${patch.address} -> ${patch.cell.raw}`)
    .join(", ")}`,
);

// The proposal only waits in the queue. The workbook is the same object as before.
console.log(`workbook unchanged: ${JSON.stringify(workbook) === before}`);
const afterQueue = await call("workbook_get_context");
const b3 = afterQueue.cells.find((cell) => cell.address === "B3");
console.log(`B3 still holds: ${b3.raw}`);

await client.close();
await server.close();
