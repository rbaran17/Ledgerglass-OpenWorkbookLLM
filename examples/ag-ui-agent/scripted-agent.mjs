/**
 * A responder with no model: fixed rules, so the example runs anywhere and its check is
 * repeatable. It shows the two things an agent does with Ledgerglass's frontend tools.
 */
import { randomUUID } from "node:crypto";

function lastUserText(input) {
  const message = input.messages.findLast((item) => item.role === "user");
  return typeof message?.content === "string" ? message.content : "";
}

function selectionFrom(input) {
  for (const item of input.context)
    try {
      const value = JSON.parse(item.value);
      if (value && Array.isArray(value.cells)) return value;
    } catch {
      // Not the selection.
    }
  return undefined;
}

export async function scriptedAgent(input) {
  const toolNames = new Set(input.tools.map((tool) => tool.name));
  const toolResult = input.messages.findLast((item) => item.role === "tool");
  if (toolResult && input.messages.at(-1) === toolResult) {
    const result = JSON.parse(toolResult.content);
    return {
      text: result.shown
        ? `The workbook is on screen with ${result.formulaErrorCount} formula error(s).`
        : `The workbook could not be shown: ${(result.problems ?? []).join(" ")}`,
    };
  }

  // In the editor's assistant panel: propose a total under the numbers of the selection.
  if (toolNames.has("propose_workbook_changes")) {
    const selection = selectionFrom(input);
    const numbers = (selection?.cells ?? []).filter(
      (cell) => typeof cell.raw === "number",
    );
    const column = numbers[0]?.address.replace(/\d+/g, "");
    const inColumn = numbers.filter(
      (cell) => cell.address.replace(/\d+/g, "") === column,
    );
    const rows = inColumn.map((cell) =>
      Number(cell.address.replace(/\D+/g, "")),
    );
    const target = (selection?.cells ?? []).find(
      (cell) =>
        cell.address === `${column}${Math.max(...rows) + 1}` &&
        (cell.raw === null || cell.raw === undefined || cell.raw === ""),
    );
    if (!/total|sum/i.test(lastUserText(input)) || !target)
      return {
        text: `The selection ${selection?.range ?? ""} has ${numbers.length} numeric cell(s). Ask for a total and leave an empty cell under them in the selection.`,
      };
    const first = `${column}${Math.min(...rows)}`;
    const last = `${column}${Math.max(...rows)}`;
    return {
      toolCalls: [
        {
          id: randomUUID(),
          name: "propose_workbook_changes",
          args: {
            explanation: `Adds a total of ${first}:${last} in ${target.address}.`,
            citations: inColumn.map((cell) => cell.address),
            patches: [
              { address: target.address, value: `=SUM(${first}:${last})` },
            ],
          },
        },
      ],
    };
  }

  // In a chat host: show a workbook.
  if (toolNames.has("render_workbook"))
    return {
      text: "Here is a small budget.",
      toolCalls: [
        {
          id: randomUUID(),
          name: "render_workbook",
          args: {
            title: "Budget",
            sheets: [
              {
                name: "Plan",
                rows: [
                  ["Item", "Cost"],
                  ["Rent", 1200],
                  ["Food", 450],
                  ["Total", "=SUM(B2:B3)"],
                ],
              },
            ],
          },
        },
      ],
    };

  return { text: "No workbook tool was offered, so there is nothing to show." };
}
