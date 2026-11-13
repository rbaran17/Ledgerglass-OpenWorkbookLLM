/**
 * A minimal AG-UI 1.0 agent endpoint: it reads a RunAgentInput, asks a "responder" what the
 * agent says, and streams the answer as AG-UI events (JSON over server-sent events).
 *
 * A responder is `async (input, { signal }) => ({ text, toolCalls })`, where `toolCalls` is a
 * list of `{ id, name, args }` (args as an object). The tools an agent may call are the frontend
 * tools the client sent in `input.tools`: the client runs them, and answers in the messages of
 * its next run. So a run that calls a tool ends there, as AG-UI 1.0 describes for frontend tools.
 */
import { randomUUID } from "node:crypto";

const maxBodyBytes = 2 * 1024 * 1024;

async function readJson(request) {
  const chunks = [];
  let size = 0;
  for await (const chunk of request) {
    size += chunk.length;
    if (size > maxBodyBytes) throw new Error("The request is too large.");
    chunks.push(chunk);
  }
  return JSON.parse(Buffer.concat(chunks).toString("utf8"));
}

function isRunInput(input) {
  return (
    input &&
    typeof input === "object" &&
    typeof input.threadId === "string" &&
    typeof input.runId === "string" &&
    Array.isArray(input.messages)
  );
}

/** Returns `async (request, response)` for a Node HTTP server. */
export function createAgUiEndpoint(respond) {
  return async function handleRun(request, response) {
    let input;
    try {
      input = await readJson(request);
      if (!isRunInput(input)) throw new Error("Not an AG-UI run input.");
    } catch (error) {
      response.writeHead(400, { "Content-Type": "application/json" });
      response.end(JSON.stringify({ error: error.message }));
      return;
    }
    response.writeHead(200, {
      "Content-Type": "text/event-stream",
      "Cache-Control": "no-store",
      Connection: "keep-alive",
    });
    const send = (event) =>
      response.write(`data: ${JSON.stringify(event)}\n\n`);
    const abort = new AbortController();
    response.on("close", () => abort.abort());
    const run = { threadId: input.threadId, runId: input.runId };
    send({ type: "RUN_STARTED", ...run, protocolVersion: "1.0" });
    try {
      const { text = "", toolCalls = [] } = await respond(
        { ...input, tools: input.tools ?? [], context: input.context ?? [] },
        { signal: abort.signal },
      );
      const messageId = randomUUID();
      if (text) {
        send({ type: "TEXT_MESSAGE_START", messageId, role: "assistant" });
        send({ type: "TEXT_MESSAGE_CONTENT", messageId, delta: text });
        send({ type: "TEXT_MESSAGE_END", messageId });
      }
      for (const call of toolCalls) {
        send({
          type: "TOOL_CALL_START",
          toolCallId: call.id,
          toolCallName: call.name,
          parentMessageId: messageId,
        });
        send({
          type: "TOOL_CALL_ARGS",
          toolCallId: call.id,
          delta: JSON.stringify(call.args),
        });
        send({ type: "TOOL_CALL_END", toolCallId: call.id });
      }
      send({ type: "RUN_FINISHED", ...run });
    } catch (error) {
      // Say what failed without passing internal details to the browser.
      console.error("Agent run failed:", error);
      send({
        type: "RUN_ERROR",
        message: "The agent could not complete this request.",
        code: "agent_error",
      });
    }
    response.end();
  };
}
