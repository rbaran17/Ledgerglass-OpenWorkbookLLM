/**
 * Checks the agent-rendered workbook against the real protocol packages: the official A2UI React
 * renderer and the AG-UI client. A small host page uses the built library the way an agent
 * application would; a local server plays the agent.
 *
 * Run `npm run build:lib` and `npm ci --prefix tests/agent-ui-host` first (`npm run test:agent-ui`
 * does both). The protocol packages live only in tests/agent-ui-host; they are not dependencies of
 * the library or the application.
 */
import assert from "node:assert/strict";
import { readFile, rm } from "node:fs/promises";
import { createServer } from "node:http";
import { build } from "esbuild";
import { chromium } from "playwright";

const repositoryRoot = process.cwd();
const outputDirectory = "work/agent-ui-host";
await rm(outputDirectory, { recursive: true, force: true });

/** One copy of React for the page: the library's, not a second one from the fixture. */
const useRepositoryReact = {
  name: "repository-react",
  setup(esbuild) {
    esbuild.onResolve({ filter: /^react(-dom)?(\/.*)?$/ }, (args) =>
      args.pluginData?.fromRepository
        ? undefined
        : esbuild.resolve(args.path, {
            resolveDir: repositoryRoot,
            kind: args.kind,
            pluginData: { fromRepository: true },
          }),
    );
  },
};
await build({
  entryPoints: ["tests/agent-ui-host/host.jsx"],
  outdir: `${outputDirectory}/assets`,
  bundle: true,
  format: "esm",
  jsx: "automatic",
  minify: true,
  define: { "process.env.NODE_ENV": '"production"' },
  plugins: [useRepositoryReact],
  logLevel: "error",
});

const budgetArguments = JSON.stringify({
  title: "Budget",
  sheets: [
    {
      name: "Plan",
      rows: [
        ["Item", "Cost"],
        ["Rent", 1200],
        ["Food", 450],
      ],
      cells: { B4: "=SUM(B2:B3)", C4: "=1/0" },
    },
  ],
});
const runInputs = [];
function sendEvents(response, events) {
  response.writeHead(200, { "content-type": "text/event-stream" });
  for (const event of events)
    response.write(`data: ${JSON.stringify(event)}\n\n`);
  response.end();
}
/** The stand-in agent: calls render_workbook first, then replies once it has the tool result. */
async function handleAgentRun(request, response) {
  let body = "";
  for await (const chunk of request) body += chunk;
  const input = JSON.parse(body);
  runInputs.push(input);
  const lifecycle = { threadId: input.threadId, runId: input.runId };
  const started = { type: "RUN_STARTED", ...lifecycle, protocolVersion: "1.0" };
  const finished = { type: "RUN_FINISHED", ...lifecycle };
  if (!input.messages.some((message) => message.role === "tool"))
    return sendEvents(response, [
      started,
      {
        type: "TOOL_CALL_START",
        toolCallId: "call-1",
        toolCallName: "render_workbook",
        parentMessageId: "msg-1",
      },
      {
        type: "TOOL_CALL_ARGS",
        toolCallId: "call-1",
        delta: budgetArguments.slice(0, 40),
      },
      {
        type: "TOOL_CALL_ARGS",
        toolCallId: "call-1",
        delta: budgetArguments.slice(40),
      },
      { type: "TOOL_CALL_END", toolCallId: "call-1" },
      finished,
    ]);
  return sendEvents(response, [
    started,
    { type: "TEXT_MESSAGE_START", messageId: "msg-2", role: "assistant" },
    { type: "TEXT_MESSAGE_CONTENT", messageId: "msg-2", delta: "Done." },
    { type: "TEXT_MESSAGE_END", messageId: "msg-2" },
    finished,
  ]);
}
const server = createServer(async (request, response) => {
  try {
    if (request.url === "/agent")
      return await handleAgentRun(request, response);
    if (request.url.startsWith("/assets/")) {
      const file = await readFile(outputDirectory + request.url.split("?")[0]);
      response.writeHead(200, {
        "content-type": request.url.includes(".css")
          ? "text/css"
          : "text/javascript",
      });
      return response.end(file);
    }
    response.writeHead(200, { "content-type": "text/html" });
    response.end(
      `<!doctype html><html lang="en"><head><meta charset="UTF-8"/><title>Agent host</title><link rel="stylesheet" href="/assets/host.css"/></head><body><div id="host"></div><script type="module" src="/assets/host.js"></script></body></html>`,
    );
  } catch (error) {
    response.writeHead(500);
    response.end(String(error));
  }
});
await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
const origin = `http://127.0.0.1:${server.address().port}`;

const browser = await chromium.launch(
  process.env.CHROME_PATH ? { executablePath: process.env.CHROME_PATH } : {},
);
async function openHost(mode) {
  const page = await browser.newPage({
    viewport: { width: 1300, height: 900 },
  });
  const errors = [];
  page.on("pageerror", (error) => errors.push(error.message));
  await page.goto(`${origin}/?mode=${mode}`);
  await page.waitForFunction(() => window.ready);
  const cell = (address) => page.locator(`[data-address="${address}"]`).first();
  const expectCell = async (address, pattern) => {
    await cell(address).waitFor();
    await assert.doesNotReject(
      page.waitForFunction(
        ([selector, source]) =>
          new RegExp(source).test(
            document.querySelector(selector)?.getAttribute("aria-label") ?? "",
          ),
        [`[data-address="${address}"]`, pattern.source],
        { timeout: 10000 },
      ),
      `${address} should show ${pattern}`,
    );
  };
  return { page, errors, cell, expectCell };
}

try {
  // A2UI: the agent places a Workbook component on a surface.
  {
    const { page, errors, cell, expectCell } = await openHost("a2ui");
    const capabilities = await page.evaluate(() => window.capabilities);
    const catalogId = await page.evaluate(() => window.workbookCatalogId);
    assert.ok(capabilities["1.0"].supportedCatalogIds.includes(catalogId));
    const catalogDocument = await page.evaluate(() => window.catalogJson);
    assert.equal(catalogDocument.catalogId, catalogId);
    assert.ok(catalogDocument.components.Workbook.properties.sheets);
    assert.match(catalogDocument.instructions, /Workbook component/);
    console.log("PASS A2UI: the renderer advertises the workbook catalog");

    const send = (messages) =>
      page.evaluate((list) => window.send(list), messages);
    const salesSheets = (first, second) => [
      {
        name: "Sales",
        rows: [
          ["Item", "Amount"],
          ["A", first],
          ["B", second],
        ],
        cells: { B4: "=SUM(B2:B3)" },
      },
    ];
    assert.equal(
      await send([
        { version: "1.0", createSurface: { surfaceId: "s1", catalogId } },
        {
          version: "1.0",
          updateDataModel: {
            surfaceId: "s1",
            path: "/",
            value: { sheets: salesSheets(10, 32) },
          },
        },
        {
          version: "1.0",
          updateComponents: {
            surfaceId: "s1",
            components: [
              {
                id: "root",
                component: "Workbook",
                title: "Q3 sales",
                sheets: { "@path": "/sheets" },
                height: 420,
                onChange: {
                  event: {
                    name: "workbook_changed",
                    context: { sheets: { "@path": "/sheets" } },
                  },
                },
              },
            ],
          },
        },
      ]),
      "ok",
    );
    await expectCell("B4", /\b42\b/);
    console.log("PASS A2UI: agent messages render a calculated workbook");

    await cell("B2").dblclick();
    await page.keyboard.press("Control+A");
    await page.keyboard.type("100");
    await page.keyboard.press("Enter");
    await expectCell("B4", /\b132\b/);
    await page.waitForFunction(() => window.log.some((entry) => entry.action));
    const dataModel = await page.evaluate(() => window.dataModel("s1"));
    assert.equal(dataModel.sheets[0].cells.B2, 100);
    const action = (await page.evaluate(() => window.log)).find(
      (entry) => entry.action,
    ).action;
    assert.equal(action.name, "workbook_changed");
    assert.equal(action.sourceComponentId, "root");
    assert.equal(action.context.sheets[0].cells.B2, 100);
    console.log(
      "PASS A2UI: the user's edit reaches the data model and the agent's action",
    );

    assert.equal(
      await send([
        {
          version: "1.0",
          updateDataModel: {
            surfaceId: "s1",
            path: "/sheets",
            value: salesSheets(1, 2),
          },
        },
      ]),
      "ok",
    );
    await expectCell("B4", /\b3\b/);
    console.log("PASS A2UI: a later agent update replaces the workbook");

    assert.equal(
      await send([
        {
          version: "1.0",
          updateDataModel: {
            surfaceId: "s1",
            path: "/sheets",
            value: [{ name: "Bad/Name" }],
          },
        },
      ]),
      "ok",
    );
    await page.getByRole("alert").waitFor();
    assert.match(
      await page.getByRole("alert").innerText(),
      /could not be shown[\s\S]*may not contain/,
    );
    console.log("PASS A2UI: a description that breaks the rules is explained");
    assert.deepEqual(errors, []);
    await page.close();
  }

  // AG-UI: the agent calls the render_workbook frontend tool.
  {
    const { page, errors, expectCell } = await openHost("agui");
    assert.deepEqual(await page.evaluate(() => window.run()), {
      ok: true,
      newMessages: 1,
    });
    assert.deepEqual(
      runInputs[0].tools.map((tool) => tool.name),
      ["render_workbook"],
    );
    assert.equal(runInputs[0].tools[0].parameters.type, "object");
    await expectCell("B4", /1,?650/);
    console.log("PASS AG-UI: a render_workbook tool call shows the workbook");

    assert.equal((await page.evaluate(() => window.run())).ok, true);
    const toolMessage = runInputs[1].messages.find(
      (message) => message.role === "tool",
    );
    assert.equal(toolMessage.toolCallId, "call-1");
    const toolResult = JSON.parse(toolMessage.content);
    assert.equal(toolResult.shown, true);
    assert.deepEqual(toolResult.formulaErrors, [
      { sheet: "Plan", address: "C4", formula: "=1/0", error: "#DIV/0!" },
    ]);
    console.log(
      "PASS AG-UI: the next run carries the tool result, including the formula error",
    );
    assert.deepEqual(errors, []);
    await page.close();
  }
} finally {
  await browser.close();
  server.close();
}

// The example agent endpoint, read by the official AG-UI client: the client checks the order and
// shape of every event, so a stream it accepts follows the protocol.
{
  const { createRequire } = await import("node:module");
  const { HttpAgent } = createRequire(
    `${repositoryRoot}/tests/agent-ui-host/host.jsx`,
  )("@ag-ui/client");
  const { startServer } = await import("../examples/ag-ui-agent/server.mjs");
  const { renderWorkbookTool, createWorkbookToolSubscriber } =
    await import("../lib/node.js");
  const agentServer = await startServer({ port: 0, agent: "scripted" });
  try {
    const agent = new HttpAgent({
      url: `http://127.0.0.1:${agentServer.address().port}/agent`,
    });
    const shown = [];
    agent.subscribe(
      createWorkbookToolSubscriber({ onWorkbook: (call) => shown.push(call) }),
    );
    agent.addMessage({ id: "u1", role: "user", content: "Show a budget" });
    await agent.runAgent({ tools: [renderWorkbookTool] });
    assert.equal(shown.length, 1);
    assert.equal(shown[0].workbook.sheets[0].name, "Plan");
    agent.addMessage(shown[0].result);
    const second = await agent.runAgent({ tools: [renderWorkbookTool] });
    assert.match(
      second.newMessages.at(-1).content,
      /on screen with 0 formula error/,
    );
    console.log(
      "PASS AG-UI: the official client accepts the example agent's event stream over two runs",
    );
  } finally {
    agentServer.close();
  }
}
