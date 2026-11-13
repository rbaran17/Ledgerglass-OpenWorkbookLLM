import { afterAll, beforeAll, describe, expect, test } from "vitest";
import { createServer } from "node:http";
import {
  PROPOSE_WORKBOOK_CHANGES_TOOL_NAME,
  createAgUiAdapter,
  proposeWorkbookChangesTool,
  runAgUiAgent,
} from "../../src/integrations/ag-ui-adapter.js";
import {
  buildAssistantSelectionContext,
  validateAssistantProposal,
} from "../../src/app/19b-assistant-protocol.js";
import { createBlankWorkbook } from "../../src/app/03-workbook-model.js";
import { createAgUiEndpoint } from "../../examples/ag-ui-agent/ag-ui-endpoint.mjs";
import { scriptedAgent } from "../../examples/ag-ui-agent/scripted-agent.mjs";

const encoder = new TextEncoder();

/** Server-sent events for AG-UI events, one `data:` block each. */
function sseText(events, eol = "\n") {
  return events
    .map((event) => `data: ${JSON.stringify(event)}${eol}${eol}`)
    .join("");
}

/** A run's events followed by RUN_FINISHED, as server-sent event text. */
function runText(...events) {
  return sseText([
    ...events,
    { type: "RUN_FINISHED", threadId: "t", runId: "r" },
  ]);
}

/** The events of one tool call, with its arguments as the given object. */
function toolCallEvents(
  args,
  { name = PROPOSE_WORKBOOK_CHANGES_TOOL_NAME, id = "call-1" } = {},
) {
  return [
    { type: "TOOL_CALL_START", toolCallId: id, toolCallName: name },
    { type: "TOOL_CALL_ARGS", toolCallId: id, delta: JSON.stringify(args) },
    { type: "TOOL_CALL_END", toolCallId: id },
  ];
}

/**
 * A response whose body is the text, delivered in pieces of `chunkSize` bytes (the whole text
 * when no size is given), so events can arrive split across chunks.
 */
function streamed(text, { chunkSize } = {}) {
  const bytes = encoder.encode(text);
  const size = chunkSize ?? bytes.length;
  return new Response(
    new ReadableStream({
      start(controller) {
        for (let at = 0; at < bytes.length; at += size)
          controller.enqueue(bytes.slice(at, at + size));
        controller.close();
      },
    }),
    { status: 200, headers: { "Content-Type": "text/event-stream" } },
  );
}

/** A run whose reply is one line of text and no tool call. */
const replyText = runText({
  type: "TEXT_MESSAGE_CONTENT",
  messageId: "m",
  delta: "Noted.",
});

/** A fake `fetch` that answers every call with `body` and records what it was asked. */
function fakeAgent(body) {
  const calls = [];
  const fetch = async (url, init) => {
    calls.push({ url, init, body: JSON.parse(init.body) });
    return typeof body === "string" ? streamed(body) : body;
  };
  return { fetch, calls };
}

const runInput = {
  protocolVersion: "1.0",
  threadId: "t",
  runId: "r",
  state: {},
  messages: [],
  tools: [],
  context: [],
  forwardedProps: {},
};

/** A 3-row budget with an empty B5, and the context for A1:B5 of it. */
function budgetFixture() {
  const workbook = createBlankWorkbook("Budget");
  const sheetId = workbook.sheets[0].id;
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
  const context = buildAssistantSelectionContext(workbook, {
    sheetId,
    range: "A1:B5",
  });
  return { sheetId, context };
}

/** A small context for the adapter tests that do not validate the proposal. */
const plainContext = {
  workbookId: "wb-1",
  revision: 3,
  sheetId: "s1",
  sheetName: "Sheet1",
  range: "A1:B5",
  cells: [{ address: "B2", raw: 1200, value: 1200 }],
};

describe("proposeWorkbookChangesTool", () => {
  test("is named propose_workbook_changes and has a description", () => {
    expect(PROPOSE_WORKBOOK_CHANGES_TOOL_NAME).toBe("propose_workbook_changes");
    expect(proposeWorkbookChangesTool.name).toBe(
      PROPOSE_WORKBOOK_CHANGES_TOOL_NAME,
    );
    expect(proposeWorkbookChangesTool.description).toMatch(
      /approved selection/,
    );
  });

  test("takes an object schema that requires explanation, citations and patches", () => {
    const { parameters } = proposeWorkbookChangesTool;
    expect(parameters.type).toBe("object");
    expect(parameters.required).toEqual([
      "explanation",
      "citations",
      "patches",
    ]);
    expect(Object.keys(parameters.properties).sort()).toEqual([
      "citations",
      "explanation",
      "patches",
    ]);
  });

  test("is frozen", () => {
    expect(Object.isFrozen(proposeWorkbookChangesTool)).toBe(true);
  });
});

describe("runAgUiAgent", () => {
  test("POSTs the input as JSON with the given headers, signal and credentials", async () => {
    const fake = fakeAgent(replyText);
    const controller = new AbortController();
    await runAgUiAgent({
      url: "https://agent.test/run",
      input: runInput,
      headers: { "X-Api-Key": "k" },
      signal: controller.signal,
      credentials: "include",
      fetch: fake.fetch,
    });
    const [call] = fake.calls;
    expect(call.url).toBe("https://agent.test/run");
    expect(call.init.method).toBe("POST");
    expect(call.init.headers).toEqual({
      "Content-Type": "application/json",
      Accept: "text/event-stream",
      "X-Api-Key": "k",
    });
    expect(call.init.body).toBe(JSON.stringify(runInput));
    expect(call.init.signal).toBe(controller.signal);
    expect(call.init.credentials).toBe("include");

    await runAgUiAgent({ url: "x", input: runInput, fetch: fake.fetch });
    expect(fake.calls[1].init.credentials).toBe("same-origin");
  });

  test("collects text from TEXT_MESSAGE_CONTENT and TEXT_MESSAGE_CHUNK deltas in order", async () => {
    const fake = fakeAgent(
      runText(
        { type: "TEXT_MESSAGE_START", messageId: "m", role: "assistant" },
        { type: "TEXT_MESSAGE_CONTENT", messageId: "m", delta: "Hello, " },
        { type: "TEXT_MESSAGE_CHUNK", messageId: "m", delta: "wor" },
        { type: "TEXT_MESSAGE_CONTENT", messageId: "m", delta: "ld." },
        { type: "TEXT_MESSAGE_END", messageId: "m" },
      ),
    );
    const result = await runAgUiAgent({
      url: "x",
      input: runInput,
      fetch: fake.fetch,
    });
    expect(result).toEqual({ text: "Hello, world.", toolCalls: [] });
  });

  test("joins tool call arguments that arrive in several ARGS events", async () => {
    const fake = fakeAgent(
      runText(
        {
          type: "TOOL_CALL_START",
          toolCallId: "c1",
          toolCallName: "propose_workbook_changes",
        },
        { type: "TOOL_CALL_ARGS", toolCallId: "c1", delta: '{"explanation":' },
        { type: "TOOL_CALL_ARGS", toolCallId: "c1", delta: '"Sum"' },
        { type: "TOOL_CALL_ARGS", toolCallId: "c1", delta: "}" },
        { type: "TOOL_CALL_END", toolCallId: "c1" },
      ),
    );
    const { toolCalls } = await runAgUiAgent({
      url: "x",
      input: runInput,
      fetch: fake.fetch,
    });
    expect(toolCalls).toEqual([
      {
        id: "c1",
        name: "propose_workbook_changes",
        args: '{"explanation":"Sum"}',
      },
    ]);
  });

  test("names a tool call from the first TOOL_CALL_CHUNK and takes later chunks without ids", async () => {
    const fake = fakeAgent(
      runText(
        {
          type: "TOOL_CALL_CHUNK",
          toolCallId: "c2",
          toolCallName: "propose_workbook_changes",
          delta: '{"explanation":',
        },
        { type: "TOOL_CALL_CHUNK", delta: '"Sum","citations":[],' },
        { type: "TOOL_CALL_CHUNK", delta: '"patches":[]}' },
      ),
    );
    const { toolCalls } = await runAgUiAgent({
      url: "x",
      input: runInput,
      fetch: fake.fetch,
    });
    expect(toolCalls).toEqual([
      {
        id: "c2",
        name: "propose_workbook_changes",
        args: '{"explanation":"Sum","citations":[],"patches":[]}',
      },
    ]);
  });

  test("returns two tool calls from one run in order", async () => {
    const fake = fakeAgent(
      runText(
        ...toolCallEvents({ a: 1 }, { name: "first", id: "c1" }),
        ...toolCallEvents({ b: 2 }, { name: "second", id: "c2" }),
      ),
    );
    const { toolCalls } = await runAgUiAgent({
      url: "x",
      input: runInput,
      fetch: fake.fetch,
    });
    expect(toolCalls).toEqual([
      { id: "c1", name: "first", args: '{"a":1}' },
      { id: "c2", name: "second", args: '{"b":2}' },
    ]);
  });

  test("ignores events of unknown or other types", async () => {
    const fake = fakeAgent(
      runText(
        { type: "RUN_STARTED", threadId: "t", runId: "r" },
        { type: "STATE_SNAPSHOT", snapshot: {} },
        { type: "CUSTOM", name: "whatever", value: 1 },
        { type: "SOMETHING_NEW", delta: "not text" },
        { type: "TEXT_MESSAGE_CONTENT", messageId: "m", delta: "Hi" },
      ),
    );
    const result = await runAgUiAgent({
      url: "x",
      input: runInput,
      fetch: fake.fetch,
    });
    expect(result).toEqual({ text: "Hi", toolCalls: [] });
  });

  test("reads the same reply when the bytes arrive one at a time", async () => {
    const text = sseText([
      { type: "TEXT_MESSAGE_CONTENT", messageId: "m", delta: "café ✓" },
      ...toolCallEvents({ patches: [] }),
      { type: "RUN_FINISHED", threadId: "t", runId: "r" },
    ]);
    const whole = await runAgUiAgent({
      url: "x",
      input: runInput,
      fetch: fakeAgent(text).fetch,
    });
    const oneByte = await runAgUiAgent({
      url: "x",
      input: runInput,
      fetch: async () => streamed(text, { chunkSize: 1 }),
    });
    expect(oneByte).toEqual(whole);
    expect(oneByte.text).toBe("café ✓");
  });

  test("joins a data field that is split over several data lines", async () => {
    const raw =
      'data: {"type":"TEXT_MESSAGE_CONTENT",\ndata: "messageId":"m",\ndata: "delta":"Hi"}\n\n' +
      'data: {"type":"RUN_FINISHED","threadId":"t","runId":"r"}\n\n';
    const { text } = await runAgUiAgent({
      url: "x",
      input: runInput,
      fetch: fakeAgent(raw).fetch,
    });
    expect(text).toBe("Hi");
  });

  test("reads a stream whose lines end in CRLF", async () => {
    const raw = sseText(
      [
        { type: "TEXT_MESSAGE_CONTENT", messageId: "m", delta: "Hi" },
        ...toolCallEvents({ patches: [] }),
        { type: "RUN_FINISHED", threadId: "t", runId: "r" },
      ],
      "\r\n",
    );
    const result = await runAgUiAgent({
      url: "x",
      input: runInput,
      fetch: fakeAgent(raw).fetch,
    });
    expect(result.text).toBe("Hi");
    expect(result.toolCalls.map((call) => call.name)).toEqual([
      "propose_workbook_changes",
    ]);
  });

  test("throws with the status when the service answers an error", async () => {
    const fake = fakeAgent(new Response("busy", { status: 503 }));
    await expect(
      runAgUiAgent({ url: "x", input: runInput, fetch: fake.fetch }),
    ).rejects.toThrow(/answered 503\. No change was proposed\./);
  });

  test("throws the agent's message when it reports RUN_ERROR", async () => {
    const fake = fakeAgent(
      runText({ type: "RUN_ERROR", message: "model quota exceeded" }),
    );
    await expect(
      runAgUiAgent({ url: "x", input: runInput, fetch: fake.fetch }),
    ).rejects.toThrow("The agent reported an error: model quota exceeded");
  });

  test("throws a generic message when RUN_ERROR has none", async () => {
    const fake = fakeAgent(runText({ type: "RUN_ERROR" }));
    await expect(
      runAgUiAgent({ url: "x", input: runInput, fetch: fake.fetch }),
    ).rejects.toThrow("The agent reported an error. No change was proposed.");
  });

  test("cuts a long RUN_ERROR message to 500 characters", async () => {
    const fake = fakeAgent(
      runText({ type: "RUN_ERROR", message: `${"x".repeat(500)}END` }),
    );
    const error = await runAgUiAgent({
      url: "x",
      input: runInput,
      fetch: fake.fetch,
    }).catch((caught) => caught);
    expect(error).toBeInstanceOf(Error);
    expect(error.message).toContain(`: ${"x".repeat(500)}`);
    expect(error.message).not.toContain("END");
  });

  test("throws when the stream ends without RUN_FINISHED", async () => {
    const fake = fakeAgent(
      sseText([{ type: "TEXT_MESSAGE_CONTENT", messageId: "m", delta: "Hi" }]),
    );
    await expect(
      runAgUiAgent({ url: "x", input: runInput, fetch: fake.fetch }),
    ).rejects.toThrow(/ended early/);
  });

  test("throws when an event is not JSON", async () => {
    const fake = fakeAgent("data: not json\n\n");
    await expect(
      runAgUiAgent({ url: "x", input: runInput, fetch: fake.fetch }),
    ).rejects.toThrow(/not valid JSON/);
  });
});

describe("createAgUiAdapter", () => {
  test("needs a url", () => {
    expect(() => createAgUiAdapter()).toThrow(TypeError);
    expect(() => createAgUiAdapter({ url: "" })).toThrow(
      "createAgUiAdapter needs a url.",
    );
  });

  test("is named Workbook assistant by default, or as given", () => {
    expect(createAgUiAdapter({ url: "x" }).name).toBe("Workbook assistant");
    expect(createAgUiAdapter({ url: "x", name: "Agent" }).name).toBe("Agent");
  });

  test("offers chat history unless chat is false", () => {
    expect(createAgUiAdapter({ url: "x" }).chat).toEqual({
      history: true,
      attachments: false,
    });
    expect(createAgUiAdapter({ url: "x", chat: false })).not.toHaveProperty(
      "chat",
    );
  });

  test("sends the AG-UI run input with the one frontend tool", async () => {
    const fake = fakeAgent(replyText);
    const adapter = createAgUiAdapter({ url: "x", fetch: fake.fetch });
    await adapter.propose("Add a total", plainContext);
    const body = fake.calls[0].body;
    expect(body.protocolVersion).toBe("1.0");
    expect(body.state).toEqual({});
    expect(body.forwardedProps).toEqual({});
    expect(body.tools).toEqual([proposeWorkbookChangesTool]);
  });

  test("keeps one thread per conversation and a fresh run each call", async () => {
    const fake = fakeAgent(replyText);
    const adapter = createAgUiAdapter({ url: "x", fetch: fake.fetch });
    const other = createAgUiAdapter({ url: "x", fetch: fake.fetch });
    const history = {
      version: 1,
      messages: [
        { role: "user", text: "one" },
        { role: "assistant", text: "reply" },
      ],
      attachments: [],
    };
    await adapter.propose("one", plainContext);
    await adapter.propose("two", plainContext, undefined, history);
    await other.propose("three", plainContext);
    const [first, second, third] = fake.calls.map((call) => call.body);
    expect(second.threadId).toBe(first.threadId);
    expect(third.threadId).not.toBe(first.threadId);
    expect(second.runId).not.toBe(first.runId);
    expect(first.runId).toEqual(expect.any(String));
  });

  test("a message keeps its id on every run that carries it", async () => {
    const fake = fakeAgent(replyText);
    const adapter = createAgUiAdapter({ url: "x", fetch: fake.fetch });
    await adapter.propose("one", plainContext);
    await adapter.propose("two", plainContext, undefined, {
      version: 1,
      messages: [
        { role: "user", text: "one" },
        { role: "assistant", text: "reply" },
      ],
      attachments: [],
    });
    const [first, second] = fake.calls.map((call) => call.body);
    expect(second.messages[0]).toEqual(first.messages[0]);
    const ids = second.messages.map((message) => message.id);
    expect(new Set(ids).size).toBe(3);
  });

  test("a run with no earlier messages after the first starts a new thread", async () => {
    const fake = fakeAgent(replyText);
    const adapter = createAgUiAdapter({ url: "x", fetch: fake.fetch });
    await adapter.propose("one", plainContext);
    await adapter.propose("a new chat", plainContext);
    const [first, second] = fake.calls.map((call) => call.body);
    expect(second.threadId).not.toBe(first.threadId);
  });

  test("sends the approved selection as context", async () => {
    const fake = fakeAgent(replyText);
    const adapter = createAgUiAdapter({ url: "x", fetch: fake.fetch });
    await adapter.propose("Add a total", plainContext);
    const [item] = fake.calls[0].body.context;
    expect(item.description).toMatch(/approved selection/);
    expect(JSON.parse(item.value)).toEqual(plainContext);
  });

  test("ends the messages with the prompt as the user's message", async () => {
    const fake = fakeAgent(replyText);
    const adapter = createAgUiAdapter({ url: "x", fetch: fake.fetch });
    await adapter.propose("Add a total", plainContext);
    const last = fake.calls[0].body.messages.at(-1);
    expect(last).toEqual({
      id: expect.any(String),
      role: "user",
      content: "Add a total",
    });
  });

  test("sends earlier chat messages before the prompt, skipping others", async () => {
    const fake = fakeAgent(replyText);
    const adapter = createAgUiAdapter({ url: "x", fetch: fake.fetch });
    await adapter.propose("Add a total", plainContext, undefined, {
      messages: [
        { role: "user", text: "What is here?" },
        { role: "assistant", text: "Three numbers." },
        { role: "system", text: "not sent" },
        { role: "assistant", text: "" },
      ],
    });
    const messages = fake.calls[0].body.messages.map(({ id, ...rest }) => {
      expect(id).toEqual(expect.any(String));
      return rest;
    });
    expect(messages).toEqual([
      { role: "user", content: "What is here?" },
      { role: "assistant", content: "Three numbers." },
      { role: "user", content: "Add a total" },
    ]);
  });

  test("sends no chat history when chat is false", async () => {
    const fake = fakeAgent(replyText);
    const adapter = createAgUiAdapter({
      url: "x",
      chat: false,
      fetch: fake.fetch,
    });
    await adapter.propose("Add a total", plainContext, undefined, {
      messages: [{ role: "user", text: "Earlier" }],
    });
    expect(fake.calls[0].body.messages.map((item) => item.content)).toEqual([
      "Add a total",
    ]);
  });

  test("turns a tool call into a proposal with the context's identifiers", async () => {
    const fake = fakeAgent(
      runText(
        ...toolCallEvents({
          explanation: "Adds a total.",
          citations: ["B2", "B3"],
          patches: [{ address: "B5", value: "=SUM(B2:B3)" }],
        }),
      ),
    );
    const adapter = createAgUiAdapter({ url: "x", fetch: fake.fetch });
    const proposal = await adapter.propose("Add a total", plainContext);
    expect(proposal).toEqual({
      id: expect.any(String),
      workbookId: "wb-1",
      revision: 3,
      explanation: "Adds a total.",
      citations: ["B2", "B3"],
      patches: [{ sheetId: "s1", address: "B5", cell: { raw: "=SUM(B2:B3)" } }],
    });
  });

  test("maps a null or missing value to an empty cell", async () => {
    const fake = fakeAgent(
      runText(
        ...toolCallEvents({
          explanation: "Clears.",
          citations: [],
          patches: [{ address: "B2", value: null }, { address: "B3" }],
        }),
      ),
    );
    const adapter = createAgUiAdapter({ url: "x", fetch: fake.fetch });
    const { patches } = await adapter.propose("Clear", plainContext);
    expect(patches.map((patch) => patch.cell)).toEqual([
      { raw: null },
      { raw: null },
    ]);
  });

  test("turns a text-only reply into a proposal with no changes", async () => {
    const fake = fakeAgent(
      runText({
        type: "TEXT_MESSAGE_CONTENT",
        messageId: "m",
        delta: "  Sales are up.  ",
      }),
    );
    const adapter = createAgUiAdapter({ url: "x", fetch: fake.fetch });
    const proposal = await adapter.propose("What changed?", plainContext);
    expect(proposal).toEqual({
      id: expect.any(String),
      workbookId: "wb-1",
      revision: 3,
      explanation: "Sales are up.",
      citations: [],
      patches: [],
    });
  });

  test("throws when the reply is empty", async () => {
    const fake = fakeAgent(runText());
    const adapter = createAgUiAdapter({ url: "x", fetch: fake.fetch });
    await expect(adapter.propose("Hi", plainContext)).rejects.toThrow(
      /sent no reply/,
    );
  });

  test("throws when the tool arguments are not JSON", async () => {
    const fake = fakeAgent(
      runText(
        {
          type: "TOOL_CALL_START",
          toolCallId: "c",
          toolCallName: "propose_workbook_changes",
        },
        { type: "TOOL_CALL_ARGS", toolCallId: "c", delta: "{bad" },
        { type: "TOOL_CALL_END", toolCallId: "c" },
      ),
    );
    const adapter = createAgUiAdapter({ url: "x", fetch: fake.fetch });
    await expect(adapter.propose("Hi", plainContext)).rejects.toThrow(
      /not valid JSON/,
    );
  });

  test("throws when the tool arguments have no list of patches", async () => {
    const fake = fakeAgent(
      runText(...toolCallEvents({ explanation: "x", citations: [] })),
    );
    const adapter = createAgUiAdapter({ url: "x", fetch: fake.fetch });
    await expect(adapter.propose("Hi", plainContext)).rejects.toThrow(
      /no list of changes/,
    );
  });

  test("explains a proposal with the reply's text when the tool gives no explanation", async () => {
    const withText = fakeAgent(
      runText(
        {
          type: "TEXT_MESSAGE_CONTENT",
          messageId: "m",
          delta: "Here is the sum.",
        },
        ...toolCallEvents({ citations: [], patches: [] }),
      ),
    );
    const withoutText = fakeAgent(
      runText(...toolCallEvents({ citations: [], patches: [] })),
    );
    const explained = await createAgUiAdapter({
      url: "x",
      fetch: withText.fetch,
    }).propose("Sum", plainContext);
    const fallback = await createAgUiAdapter({
      url: "x",
      fetch: withoutText.fetch,
    }).propose("Sum", plainContext);
    expect(explained.explanation).toBe("Here is the sum.");
    expect(fallback.explanation).toBe("Proposed changes.");
  });

  test("ignores a tool call with another name", async () => {
    const fake = fakeAgent(
      runText(
        { type: "TEXT_MESSAGE_CONTENT", messageId: "m", delta: "Hello." },
        ...toolCallEvents({ patches: [] }, { name: "render_workbook" }),
      ),
    );
    const adapter = createAgUiAdapter({ url: "x", fetch: fake.fetch });
    const proposal = await adapter.propose("Hi", plainContext);
    expect(proposal.explanation).toBe("Hello.");
    expect(proposal.patches).toEqual([]);
  });

  test("a proposal that the editor validates against the approved selection is accepted", async () => {
    const { sheetId, context } = budgetFixture();
    const fake = fakeAgent(
      runText(
        ...toolCallEvents({
          explanation: "Adds a total of B2:B3.",
          citations: ["B2", "B3"],
          patches: [{ address: "B5", value: "=SUM(B2:B3)" }],
        }),
      ),
    );
    const adapter = createAgUiAdapter({ url: "x", fetch: fake.fetch });
    const proposal = await adapter.propose("Add a total", context);
    const checked = validateAssistantProposal(proposal, context);
    expect(checked.patches).toEqual([
      { sheetId, address: "B5", cell: { raw: "=SUM(B2:B3)" } },
    ]);
  });

  test("the adapter passes a proposal outside the selection on, and the editor refuses it", async () => {
    const { context } = budgetFixture();
    const fake = fakeAgent(
      runText(
        ...toolCallEvents({
          explanation: "Writes outside.",
          citations: [],
          patches: [{ address: "D9", value: 1 }],
        }),
      ),
    );
    const adapter = createAgUiAdapter({ url: "x", fetch: fake.fetch });
    const proposal = await adapter.propose("Write", context);
    expect(proposal.patches[0].address).toBe("D9");
    expect(() => validateAssistantProposal(proposal, context)).toThrow(
      /approved selection/,
    );
  });
});

describe("createAgUiAdapter against the example agent", () => {
  let server;
  let url;

  beforeAll(async () => {
    server = createServer(createAgUiEndpoint(scriptedAgent));
    await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
    url = `http://127.0.0.1:${server.address().port}/agent`;
  });

  afterAll(async () => {
    await new Promise((resolve) => server.close(resolve));
  });

  test("an 'Add a total' request yields a proposal that the editor accepts", async () => {
    const { sheetId, context } = budgetFixture();
    const adapter = createAgUiAdapter({ url });
    const proposal = await adapter.propose("Add a total", context);
    expect(proposal.explanation).toBe("Adds a total of B2:B4 in B5.");
    expect(proposal.citations).toEqual(["B2", "B3", "B4"]);
    const checked = validateAssistantProposal(proposal, context);
    expect(checked.patches).toEqual([
      { sheetId, address: "B5", cell: { raw: "=SUM(B2:B4)" } },
    ]);
  });

  test("a question yields a text answer and no changes", async () => {
    const { context } = budgetFixture();
    const adapter = createAgUiAdapter({ url });
    const proposal = validateAssistantProposal(
      await adapter.propose("What is in this selection?", context),
      context,
    );
    expect(proposal.explanation).toMatch(/3 numeric cell/);
    expect(proposal.patches).toEqual([]);
  });

  test("a body that is not a run input is refused with 400", async () => {
    const response = await fetch(url, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ hello: "world" }),
    });
    expect(response.status).toBe(400);
    expect(await response.json()).toEqual({
      error: "Not an AG-UI run input.",
    });
  });
});
