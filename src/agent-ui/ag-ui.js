/**
 * AG-UI: the workbook as a frontend tool. The host lists `renderWorkbookTool` in
 * `RunAgentInput.tools`; when the agent calls it, the host shows `AgentWorkbook` with the call's
 * arguments and answers with `createWorkbookToolResult` in the messages of its next run.
 *
 * Written against AG-UI 1.0. Nothing here imports the AG-UI packages: the objects follow the
 * protocol's shapes, so they work with `@ag-ui/client` or with a client of your own.
 */
import {
  WorkbookSpecError,
  buildWorkbookFromSpec,
  summarizeWorkbookForAgent,
  workbookSpecSchema,
} from "./workbook-spec.js";

export const RENDER_WORKBOOK_TOOL_NAME = "render_workbook";

/** The tool definition for `RunAgentInput.tools`. */
export const renderWorkbookTool = Object.freeze({
  name: RENDER_WORKBOOK_TOOL_NAME,
  description:
    "Show the user an interactive spreadsheet with the given sheets, values and formulas. " +
    "Use it when tabular figures, calculations or a schedule are easier to read or adjust in a grid " +
    "than in text. Formulas are calculated in the browser (Excel syntax, A1 references). " +
    "The result reports what was shown and any formula that produced an error.",
  parameters: workbookSpecSchema,
});

/**
 * Reads the arguments of a `render_workbook` call (the JSON text assembled from TOOL_CALL_ARGS, or
 * an object already parsed).
 * @returns `{ ok: true, spec, workbook, title, readOnly }` or `{ ok: false, problems }`
 */
export function readRenderWorkbookCall(toolCallArgs, options = {}) {
  let spec = toolCallArgs;
  if (typeof toolCallArgs === "string")
    try {
      spec = JSON.parse(toolCallArgs);
    } catch {
      return {
        ok: false,
        problems: ["The tool arguments are not valid JSON."],
      };
    }
  try {
    return {
      ok: true,
      spec,
      ...buildWorkbookFromSpec(spec, { id: options.toolCallId }),
    };
  } catch (error) {
    if (error instanceof WorkbookSpecError)
      return { ok: false, problems: error.problems };
    throw error;
  }
}

function createMessageId() {
  return globalThis.crypto?.randomUUID
    ? globalThis.crypto.randomUUID()
    : `tool-result-${Date.now().toString(36)}`;
}

/**
 * The tool message that answers a `render_workbook` call. Add it to the `messages` of the next
 * run. A rejected description is answered with `error` set, so the agent can correct it.
 */
export function createWorkbookToolResult(
  toolCallId,
  toolCallArgs,
  options = {},
) {
  const call = readRenderWorkbookCall(toolCallArgs, { toolCallId });
  const message = {
    id: options.messageId ?? createMessageId(),
    role: "tool",
    toolCallId,
  };
  if (!call.ok)
    return {
      ...message,
      content: JSON.stringify({ shown: false, problems: call.problems }),
      error: call.problems.join(" "),
    };
  return {
    ...message,
    content: JSON.stringify({
      shown: true,
      ...summarizeWorkbookForAgent(call.workbook),
    }),
  };
}

/**
 * An AG-UI subscriber (for `agent.subscribe` or `agent.runAgent(parameters, subscriber)`) that
 * reports each finished `render_workbook` call.
 *
 * @param handlers.onWorkbook  `({ toolCallId, spec, workbook, title, readOnly, result })`;
 *   `result` is the tool message to send back.
 * @param handlers.onInvalid   `({ toolCallId, problems, result })`
 */
export function createWorkbookToolSubscriber(handlers = {}) {
  return {
    onToolCallEndEvent({ event, toolCallName, toolCallArgs }) {
      if (toolCallName !== RENDER_WORKBOOK_TOOL_NAME) return;
      const toolCallId = event.toolCallId;
      const call = readRenderWorkbookCall(toolCallArgs, { toolCallId });
      const result = createWorkbookToolResult(toolCallId, toolCallArgs);
      if (call.ok) handlers.onWorkbook?.({ toolCallId, ...call, result });
      else
        handlers.onInvalid?.({ toolCallId, problems: call.problems, result });
    },
  };
}
