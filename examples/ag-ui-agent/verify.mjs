/**
 * Drives the editor's assistant panel against this example's AG-UI agent (the scripted one, so
 * the check needs no model and gives the same result every time).
 *
 *   node examples/ag-ui-agent/verify.mjs      (after `npm run build:lib`)
 */
import assert from "node:assert/strict";
import { chromium } from "playwright";
import { buildPage } from "./build.mjs";
import { startServer } from "./server.mjs";

await buildPage();
const server = await startServer({ port: 0, agent: "scripted" });
const origin = `http://127.0.0.1:${server.address().port}`;
const browser = await chromium.launch(
  process.env.CHROME_PATH ? { executablePath: process.env.CHROME_PATH } : {},
);

try {
  const page = await browser.newPage({
    viewport: { width: 1500, height: 950 },
  });
  const pageErrors = [];
  page.on("pageerror", (error) => pageErrors.push(error.message));
  const runs = [];
  page.on("request", (request) => {
    if (request.url().endsWith("/agent"))
      runs.push(JSON.parse(request.postData()));
  });
  const cell = (address) => page.locator(`[data-address="${address}"]`).first();
  const cellText = async (address) =>
    (await cell(address).getAttribute("aria-label")) ?? "";

  await page.goto(origin);
  await cell("A1").waitFor();

  // Approve A1:B5 as the selection the agent may see: drag from A1 to B5.
  await cell("A1").click();
  await cell("B5").click({ modifiers: ["Shift"] });
  await page.getByRole("button", { name: "AI assistant" }).click();
  const prompt = page.getByLabel("Assistant prompt");
  const conversation = page.getByLabel("Conversation");

  // 1. A question: the agent answers in text and proposes nothing.
  await prompt.fill("What is in this selection?");
  await prompt.press("Enter");
  await conversation.getByText(/3 numeric cell/).waitFor({ timeout: 30000 });
  assert.equal(runs.length, 1);
  const first = runs[0];
  assert.equal(first.protocolVersion, "1.0");
  assert.deepEqual(
    first.tools.map((tool) => tool.name),
    ["propose_workbook_changes"],
  );
  const selection = JSON.parse(first.context[0].value);
  assert.equal(selection.range, "A1:B5");
  assert.deepEqual(
    selection.cells.filter((item) => item.address === "B2")[0].raw,
    1200,
  );
  assert.equal(first.messages.at(-1).content, "What is in this selection?");
  console.log(
    "PASS a question reaches the agent with only the approved selection, and its text answer is shown",
  );

  // 2. A request for a change: the agent calls the tool; the editor shows a proposal to review.
  await prompt.fill("Add a total");
  await prompt.press("Enter");
  const apply = page.getByRole("button", { name: "Apply 1 changes" });
  await apply.waitFor({ timeout: 30000 });
  assert.match(await conversation.innerText(), /Adds a total of B2:B4 in B5/);
  assert.doesNotMatch(await cellText("B5"), /1,?770/);
  assert.equal(runs.length, 2);
  assert.ok(
    runs[1].messages.some((message) => message.role === "assistant"),
    "the earlier turns are sent as history",
  );
  assert.equal(runs[1].threadId, first.threadId);
  console.log(
    "PASS a requested change comes back as a proposal; nothing changes before review",
  );

  // 3. Accepting applies it.
  await apply.click();
  await page.waitForFunction(
    () =>
      /1,?770/.test(
        document
          .querySelector('[data-address="B5"]')
          ?.getAttribute("aria-label") ?? "",
      ),
    null,
    { timeout: 10000 },
  );
  console.log("PASS accepting the proposal puts =SUM(B2:B4) in B5 (1770)");

  assert.deepEqual(pageErrors, []);
  console.log("PASS no page errors");
} finally {
  await browser.close();
  server.close();
}
