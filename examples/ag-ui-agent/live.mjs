/**
 * A trial of the agent integration against a live model. It starts this example's server with
 * the Claude agent and runs the two flows from Node, without a browser:
 *   1. the assistant adapter: a question, then a requested change, checked by the editor's own
 *      proposal validation and applied to the workbook;
 *   2. the render_workbook tool: the agent builds a workbook, gets the result (including any
 *      formula errors) and may correct it.
 * It calls a real model, so it costs money and its output varies. It is not part of CI.
 *
 *   node examples/ag-ui-agent/live.mjs          (needs Anthropic credentials; after build:lib)
 */
import { randomUUID } from "node:crypto";
import {
  applyCellPatchTransaction,
  buildAssistantSelectionContext,
  calculateWorkbookValues,
  createAgUiAdapter,
  createBlankWorkbook,
  createWorkbookToolResult,
  readRenderWorkbookCall,
  renderWorkbookTool,
  runAgUiAgent,
  validateAssistantProposal,
} from "@ledgerglass/workbook/node";
import { startServer } from "./server.mjs";

const server = await startServer({ port: 0, agent: "claude" });
const url = `http://127.0.0.1:${server.address().port}/agent`;
const show = (label, value) =>
  console.log(
    `${label}: ${typeof value === "string" ? value : JSON.stringify(value)}`,
  );

try {
  // ---- 1. The assistant adapter -------------------------------------------------------
  console.log("== Assistant adapter (propose_workbook_changes) ==");
  let workbook = createBlankWorkbook("Budget");
  Object.assign(workbook.sheets[0].cells, {
    A1: { raw: "Item" },
    B1: { raw: "Cost" },
    A2: { raw: "Rent" },
    B2: { raw: 1200 },
    A3: { raw: "Food" },
    B3: { raw: 450 },
    A4: { raw: "Travel" },
    B4: { raw: 120 },
    A5: { raw: "Total" },
  });
  const sheetId = workbook.sheets[0].id;
  const selection = { sheetId, range: "A1:B5" };
  const adapter = createAgUiAdapter({ url });

  let context = buildAssistantSelectionContext(workbook, selection);
  const answer = validateAssistantProposal(
    await adapter.propose("Which item costs the most? Answer only.", context),
    context,
  );
  show("question -> changes proposed", answer.patches.length);
  show("question -> answer", answer.explanation);

  const proposal = validateAssistantProposal(
    await adapter.propose(
      "Put a total of the costs in B5.",
      context,
      undefined,
      {
        version: 1,
        messages: [
          { role: "user", text: "Which item costs the most? Answer only." },
          { role: "assistant", text: answer.explanation },
        ],
        attachments: [],
      },
    ),
    context,
  );
  show("change -> explanation", proposal.explanation);
  show("change -> citations", proposal.citations);
  show(
    "change -> patches",
    proposal.patches.map((patch) => [patch.address, patch.cell.raw]),
  );
  console.log("change -> passed the editor's proposal validation");
  workbook = applyCellPatchTransaction(
    workbook,
    proposal.patches,
    "Reviewer",
    "Accepted assistant proposal",
  );
  show("B5 after applying", calculateWorkbookValues(workbook).sheets.Sheet1.B5);

  // ---- 2. The render_workbook tool ----------------------------------------------------
  console.log("\n== Agent-rendered workbook (render_workbook) ==");
  const threadId = randomUUID();
  const messages = [
    {
      id: randomUUID(),
      role: "user",
      content:
        "Show me a workbook for a 250,000 loan at 6% a year repaid in 3 equal annual payments: one row per year with opening balance, payment, interest, principal and closing balance, plus totals. Use formulas.",
    },
  ];
  for (let round = 1; round <= 4; round++) {
    const { text, toolCalls } = await runAgUiAgent({
      url,
      input: {
        protocolVersion: "1.0",
        threadId,
        runId: randomUUID(),
        state: {},
        messages,
        tools: [renderWorkbookTool],
        context: [],
        forwardedProps: {},
      },
    });
    if (text) show(`round ${round} text`, text);
    const call = toolCalls.find(
      (item) => item.name === renderWorkbookTool.name,
    );
    if (!call) break;
    messages.push({
      id: randomUUID(),
      role: "assistant",
      content: text,
      toolCalls: [
        {
          id: call.id,
          type: "function",
          function: { name: call.name, arguments: call.args },
        },
      ],
    });
    const read = readRenderWorkbookCall(call.args, { toolCallId: call.id });
    const result = createWorkbookToolResult(call.id, call.args);
    messages.push(result);
    if (!read.ok) {
      show(`round ${round} rejected`, read.problems);
      continue;
    }
    const summary = JSON.parse(result.content);
    show(`round ${round} shown`, {
      title: summary.title,
      sheets: summary.sheets,
      formulaErrorCount: summary.formulaErrorCount,
    });
    if (summary.formulaErrors.length)
      show(`round ${round} formula errors`, summary.formulaErrors);
    const values = calculateWorkbookValues(read.workbook).sheets;
    const firstSheet = Object.keys(values)[0];
    show(
      `round ${round} cells of "${firstSheet}"`,
      Object.fromEntries(
        Object.entries(values[firstSheet]).map(([address, cellValue]) => [
          address,
          typeof cellValue.value === "number"
            ? Math.round(cellValue.value * 100) / 100
            : cellValue.value,
        ]),
      ),
    );
  }
} finally {
  server.close();
}
