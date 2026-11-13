/**
 * A responder backed by Claude through the Anthropic SDK. It offers Claude the frontend tools
 * the client sent, and passes Claude's tool calls back for the client to run.
 *
 * Credentials come from the environment (ANTHROPIC_API_KEY, or an `ant auth login` profile);
 * they stay on this server and never reach the browser.
 *
 * The conversation is kept per AG-UI thread, in memory, in the SDK's own message format: Claude's
 * replies have to be sent back to it exactly as they arrived, thinking blocks included, and those
 * do not survive a round trip through AG-UI messages.
 */
import Anthropic from "@anthropic-ai/sdk";

const systemPrompt = `You are the assistant of Ledgerglass, a spreadsheet for audit and finance work.

You act only through the tools you are given. They run in the user's browser:
- render_workbook shows the user an interactive spreadsheet. Use it when figures, a calculation or a schedule are easier to read or adjust in a grid than in text. Put calculations in formulas (Excel syntax, A1 references) so the user can change the inputs. Its result tells you what was shown and lists any formula that produced an error; if there are errors, call it again with them corrected.
- propose_workbook_changes proposes cell changes for the user to review. Nothing changes until the user accepts. You can read, cite and change only the cells of the approved selection given in the context, and a formula may refer only to those cells.

When the user asks a question that needs no change and no workbook, answer in a sentence or two of plain text and call no tool.
Never invent figures the user did not give you: ask, or label an assumption as an assumption.`;

const threads = new Map();
const maxThreads = 200;

function toolsFor(input) {
  return input.tools.map((tool) => ({
    name: tool.name,
    description: tool.description,
    input_schema: tool.parameters ?? { type: "object", properties: {} },
  }));
}

/**
 * Adds the AG-UI messages this thread has not seen yet, in the SDK's message format.
 * Every tool call Claude made must be answered before anything else is said, so the results
 * (or a note that none was reported) go first, in one message.
 */
function appendNewMessages(thread, input) {
  const fresh = input.messages.filter(
    (message) => !thread.seen.has(message.id),
  );
  for (const message of fresh) thread.seen.add(message.id);

  if (thread.pending.size) {
    const answers = new Map(
      fresh
        .filter((message) => message.role === "tool")
        .map((message) => [message.toolCallId, message]),
    );
    thread.messages.push({
      role: "user",
      content: [...thread.pending].map((toolCallId) => {
        const answer = answers.get(toolCallId);
        return {
          type: "tool_result",
          tool_use_id: toolCallId,
          // The assistant panel shows a proposal to the user and reports nothing back.
          content: answer
            ? String(answer.content ?? "")
            : "This was shown to the user. The application reported no further result.",
          ...(answer?.error ? { is_error: true } : {}),
        };
      }),
    });
    thread.pending.clear();
  }
  // Assistant messages are not copied: this thread already holds Claude's own replies.
  for (const message of fresh)
    if (message.role === "user" && typeof message.content === "string")
      thread.messages.push({ role: "user", content: message.content });
}

export function createClaudeAgent({
  client = new Anthropic(),
  model = "claude-opus-5-5",
} = {}) {
  return async function claudeAgent(input, { signal } = {}) {
    if (!threads.has(input.threadId)) {
      if (threads.size >= maxThreads)
        threads.delete(threads.keys().next().value);
      threads.set(input.threadId, {
        messages: [],
        seen: new Set(),
        pending: new Set(),
      });
    }
    const thread = threads.get(input.threadId);
    appendNewMessages(thread, input);

    const contextText = input.context
      .map((item) => `${item.description}\n${item.value}`)
      .join("\n\n");
    const response = await client.beta.messages.create(
      {
        model,
        max_tokens: 16000,
        // If the model declines a request on policy grounds, let the API retry it on the
        // fallback model it recommends instead of failing the turn.
        betas: ["server-side-fallback-2026-07-01"],
        fallbacks: "default",
        system: contextText
          ? [
              { type: "text", text: systemPrompt },
              { type: "text", text: contextText },
            ]
          : systemPrompt,
        tools: toolsFor(input),
        messages: thread.messages,
      },
      { signal },
    );
    if (response.stop_reason === "refusal")
      return { text: "I can't help with that request." };

    // Keep the reply exactly as it arrived, for the next request on this thread.
    thread.messages.push({ role: "assistant", content: response.content });
    const toolCalls = [];
    let text = "";
    for (const block of response.content) {
      if (block.type === "text") text += block.text;
      else if (block.type === "tool_use") {
        toolCalls.push({ id: block.id, name: block.name, args: block.input });
        thread.pending.add(block.id);
      }
    }
    return {
      text,
      toolCalls,
      usage: response.usage,
      model: response.model,
      stopReason: response.stop_reason,
    };
  };
}
