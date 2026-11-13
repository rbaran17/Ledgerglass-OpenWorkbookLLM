/**
 * Assistant adapter that talks to an AG-UI agent. The editor's assistant panel calls
 * `propose(prompt, context, signal, request)`; this adapter starts an AG-UI run with the approved
 * selection as context and one frontend tool, `propose_workbook_changes`, and turns the agent's
 * tool call into the proposal the editor reviews. An agent that only answers in text gives a
 * proposal with no changes. The editor validates the result as it does for every adapter.
 *
 * Written against AG-UI 1.0 (JSON events over server-sent events). It carries its own small
 * client, so it needs no AG-UI package.
 */

export const PROPOSE_WORKBOOK_CHANGES_TOOL_NAME = "propose_workbook_changes";

/** The frontend tool the adapter offers the agent on every run. */
export const proposeWorkbookChangesTool = Object.freeze({
  name: PROPOSE_WORKBOOK_CHANGES_TOOL_NAME,
  description:
    "Propose changes to cells of the approved selection for the user to review. Nothing is " +
    "applied until the user accepts. Only cells inside the approved selection may be changed or " +
    "cited. Call it at most once per reply, with every change in one call. To answer a question " +
    "without changing anything, reply in text and do not call this tool.",
  parameters: {
    type: "object",
    properties: {
      explanation: {
        type: "string",
        description:
          "What the changes do and why, in plain words for the reviewer.",
      },
      citations: {
        type: "array",
        description:
          'Addresses in the approved selection that support the proposal, such as ["B2", "B3"].',
        items: { type: "string" },
      },
      patches: {
        type: "array",
        description: "The cell changes.",
        items: {
          type: "object",
          properties: {
            address: {
              type: "string",
              description:
                'Cell address in the approved selection, such as "B5".',
            },
            value: {
              type: ["string", "number", "boolean", "null"],
              description:
                'New content: a number, text, true/false, or null to clear. Text starting with "=" is a formula; a formula may only refer to cells in the approved selection.',
            },
          },
          required: ["address", "value"],
          additionalProperties: false,
        },
      },
    },
    required: ["explanation", "citations", "patches"],
    additionalProperties: false,
  },
});

const selectionContextDescription =
  "The approved selection of the user's workbook, as JSON: workbookId, revision, sheetId, " +
  "sheetName, range and cells (address, raw content, calculated value). This is everything you " +
  "can see. Only these cells may be cited or changed.";

function newId() {
  return globalThis.crypto.randomUUID();
}

/** Reads a server-sent event stream and yields each event's parsed JSON data. */
async function* readEvents(response) {
  const reader = response.body.getReader();
  const decoder = new TextDecoder();
  let buffer = "";
  try {
    for (;;) {
      const { value, done } = await reader.read();
      buffer += decoder.decode(value ?? new Uint8Array(), { stream: !done });
      let boundary;
      while ((boundary = buffer.search(/\r?\n\r?\n/)) !== -1) {
        const block = buffer.slice(0, boundary);
        buffer = buffer.slice(boundary).replace(/^\r?\n\r?\n/, "");
        const data = block
          .split(/\r?\n/)
          .filter((line) => line.startsWith("data:"))
          .map((line) => line.slice(5).replace(/^ /, ""))
          .join("\n");
        if (!data) continue;
        let event;
        try {
          event = JSON.parse(data);
        } catch {
          throw new Error("The agent sent an event that is not valid JSON.");
        }
        yield event;
      }
      if (done) return;
    }
  } finally {
    reader.cancel().catch(() => {});
  }
}

/**
 * Runs one AG-UI run and collects what the agent said and which frontend tools it called.
 * @returns `{ text, toolCalls: [{ id, name, args }] }`, with `args` as the JSON text received.
 */
export async function runAgUiAgent({
  url,
  input,
  headers = {},
  signal,
  fetch: fetchImplementation = globalThis.fetch,
  credentials = "same-origin",
}) {
  const response = await fetchImplementation(url, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Accept: "text/event-stream",
      ...headers,
    },
    body: JSON.stringify(input),
    signal,
    credentials,
  });
  if (!response.ok || !response.body)
    throw new Error(
      `The agent service answered ${response.status}. No change was proposed.`,
    );
  let text = "";
  const toolCalls = new Map();
  const toolCall = (id, name) => {
    if (!toolCalls.has(id))
      toolCalls.set(id, { id, name: name ?? "", args: "" });
    const call = toolCalls.get(id);
    if (name && !call.name) call.name = name;
    return call;
  };
  let lastToolCallId;
  let finished = false;
  for await (const event of readEvents(response)) {
    switch (event.type) {
      case "TEXT_MESSAGE_CONTENT":
      case "TEXT_MESSAGE_CHUNK":
        if (typeof event.delta === "string") text += event.delta;
        break;
      case "TOOL_CALL_START":
        lastToolCallId = event.toolCallId;
        toolCall(event.toolCallId, event.toolCallName);
        break;
      case "TOOL_CALL_ARGS":
        toolCall(event.toolCallId).args += event.delta ?? "";
        break;
      case "TOOL_CALL_CHUNK": {
        // The first chunk of a call names it; later chunks may omit the identifiers.
        const id = event.toolCallId ?? lastToolCallId;
        if (!id) break;
        lastToolCallId = id;
        toolCall(id, event.toolCallName).args += event.delta ?? "";
        break;
      }
      case "RUN_ERROR":
        throw new Error(
          typeof event.message === "string" && event.message
            ? `The agent reported an error: ${event.message.slice(0, 500)}`
            : "The agent reported an error. No change was proposed.",
        );
      case "RUN_FINISHED":
        finished = true;
        break;
      default:
        break;
    }
  }
  if (!finished)
    throw new Error("The agent's reply ended early. No change was proposed.");
  return { text, toolCalls: [...toolCalls.values()] };
}

/**
 * @param options.url      The AG-UI agent endpoint (POST, replies with server-sent events).
 * @param options.headers  Extra request headers. Keep model provider keys on the agent's server.
 * @param options.name     Shown in the assistant panel. Default "Workbook assistant".
 * @param options.chat     Send earlier messages of the chat with each run. Default true.
 * @param options.fetch, options.credentials  As for `createHttpAdapter`.
 */
export function createAgUiAdapter({
  url,
  headers = {},
  name = "Workbook assistant",
  chat = true,
  fetch: fetchImplementation = globalThis.fetch,
  credentials = "same-origin",
} = {}) {
  if (typeof url !== "string" || !url)
    throw new TypeError("createAgUiAdapter needs a url.");
  // One AG-UI thread per adapter, so an agent that keeps state per thread sees one conversation.
  const threadId = newId();

  return {
    name,
    ...(chat ? { chat: { history: true, attachments: false } } : {}),
    async propose(prompt, context, signal, request) {
      const history = (chat ? (request?.messages ?? []) : [])
        .filter(
          (message) =>
            (message.role === "user" || message.role === "assistant") &&
            typeof message.text === "string" &&
            message.text,
        )
        .map((message) => ({
          id: message.id ?? newId(),
          role: message.role,
          content: message.text,
        }));
      const input = {
        protocolVersion: "1.0",
        threadId,
        runId: newId(),
        state: {},
        messages: [...history, { id: newId(), role: "user", content: prompt }],
        tools: [proposeWorkbookChangesTool],
        context: [
          {
            description: selectionContextDescription,
            value: JSON.stringify(context),
          },
        ],
        forwardedProps: {},
      };
      const { text, toolCalls } = await runAgUiAgent({
        url,
        input,
        headers,
        signal,
        fetch: fetchImplementation,
        credentials,
      });
      const call = toolCalls.find(
        (candidate) => candidate.name === PROPOSE_WORKBOOK_CHANGES_TOOL_NAME,
      );
      const base = {
        id: newId(),
        workbookId: context.workbookId,
        revision: context.revision,
      };
      if (!call) {
        if (!text.trim())
          throw new Error("The agent sent no reply. No change was proposed.");
        return {
          ...base,
          explanation: text.trim(),
          citations: [],
          patches: [],
        };
      }
      let args;
      try {
        args = JSON.parse(call.args);
      } catch {
        throw new Error(
          "The agent's proposal is not valid JSON. No change was proposed.",
        );
      }
      if (!args || typeof args !== "object" || !Array.isArray(args.patches))
        throw new Error(
          "The agent's proposal has no list of changes. No change was proposed.",
        );
      return {
        ...base,
        explanation:
          typeof args.explanation === "string" && args.explanation.trim()
            ? args.explanation.trim()
            : text.trim() || "Proposed changes.",
        citations: Array.isArray(args.citations) ? args.citations : [],
        patches: args.patches.map((patch) => ({
          sheetId: context.sheetId,
          address: patch?.address,
          cell: { raw: patch?.value ?? null },
        })),
      };
    },
  };
}
