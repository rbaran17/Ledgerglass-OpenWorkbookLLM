// Checks the host page the way a user would use it: builds public/, starts server.mjs on a free
// port, opens the page in Chromium and drives the toolbar and the grid. It also checks the
// assistant endpoint with a context built from the downloaded workbook.
// Run from the repository root:
//   CHROME_PATH=/opt/pw-browsers/chromium-1243/chrome-linux64/chrome node examples/react-host/verify.mjs
// Without CHROME_PATH, Playwright uses the browser it installed.
import assert from "node:assert/strict";
import { mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { chromium } from "playwright";
import {
  buildAssistantSelectionContext,
  calculateWorkbookValues,
  readWorkbookFile,
  validateAssistantProposal,
} from "@ledgerglass/workbook/node";
import { buildHost } from "./build.mjs";
import { startServer } from "./server.mjs";

await buildHost();
const server = await startServer({ port: 0 });
const baseUrl = `http://127.0.0.1:${server.address().port}`;
const workDir = await mkdtemp(path.join(tmpdir(), "ledgerglass-host-"));
const browser = await chromium.launch(
  process.env.CHROME_PATH ? { executablePath: process.env.CHROME_PATH } : {},
);

// The text a cell displays. Its aria-label reads "B2 1200", so the address comes off the front.
async function cellText(page, address) {
  const label = await page
    .locator(`[data-address="${address}"]`)
    .first()
    .getAttribute("aria-label");
  return (label ?? "").replace(/^[A-Z]+\d+\s?/, "").trim();
}

async function waitForCell(page, address, expected) {
  await page.waitForFunction(
    ([selector, text]) =>
      document
        .querySelector(selector)
        ?.getAttribute("aria-label")
        ?.includes(text),
    [`[data-address="${address}"]`, expected],
  );
}

// Waits for the toolbar's status line to say what the last action did.
async function waitForStatus(page, pattern) {
  await page.waitForFunction(
    (source) =>
      new RegExp(source).test(
        document.querySelector('[data-testid="status"]')?.textContent ?? "",
      ),
    pattern.source,
  );
  return page.getByTestId("status").textContent();
}

async function revision(page) {
  const text = await page.getByTestId("revision").textContent();
  return Number(/Revision (\d+)/.exec(text)[1]);
}

try {
  const page = await browser.newPage({
    viewport: { width: 1280, height: 800 },
  });
  const errors = [];
  page.on("pageerror", (error) => errors.push(error.message));
  page.on("console", (message) => {
    if (message.type() === "error") errors.push(message.text());
  });
  await page.goto(`${baseUrl}/`);
  await page
    .locator('[data-address="E2"]')
    .first()
    .waitFor({ timeout: 30_000 });

  // Starting values. The SUM shows #PENDING! until the calculation finishes, so wait for it.
  await waitForCell(page, "E2", "1770");
  assert.equal(await cellText(page, "A2"), "Rent");
  assert.equal(await cellText(page, "B2"), "1200");
  assert.equal(await cellText(page, "B3"), "450");
  assert.equal(await cellText(page, "B4"), "120");
  assert.equal(await cellText(page, "E2"), "1770");
  assert.equal(
    await page.getByTestId("workbook-name").textContent(),
    "Household budget",
  );
  const box = await page.locator("main.editor").boundingBox();
  assert.ok(box && box.height > 300, "the editor fills its container");
  console.log(
    "PASS the grid shows the starting values and the editor fills its container",
  );

  // Editing changes the SUM and the revision
  const revisionBefore = await revision(page);
  await page.locator('[data-address="B2"]').first().dblclick();
  await page.keyboard.press("Control+A");
  await page.keyboard.type("2000");
  await page.keyboard.press("Enter");
  await waitForCell(page, "E2", "2570");
  assert.equal(await cellText(page, "B2"), "2000");
  const revisionAfter = await revision(page);
  assert.ok(
    revisionAfter > revisionBefore,
    "the revision goes up after an edit",
  );
  console.log(
    `PASS typing 2000 in B2 changes the SUM to 2570 and the revision from ${revisionBefore} to ${revisionAfter}`,
  );

  // Download XLSX gives a file the library reads back with the edit
  const [download] = await Promise.all([
    page.waitForEvent("download"),
    page.getByRole("button", { name: "Download XLSX" }).click(),
  ]);
  const downloaded = path.join(workDir, download.suggestedFilename());
  await download.saveAs(downloaded);
  const { workbook, warnings } = await readWorkbookFile(downloaded);
  const values = calculateWorkbookValues(workbook).sheets.Budget;
  assert.deepEqual(values.B2, { type: "number", value: 2000 });
  assert.deepEqual(values.E2, { type: "number", value: 2570 });
  assert.equal(workbook.name, "Household budget");
  assert.ok(
    warnings.length > 0,
    "the XLSX export reports what it could not keep",
  );
  console.log(
    "PASS Download XLSX produces a file that readWorkbookFile reads back with B2 = 2000 and SUM = 2570",
  );

  // Open file: the XLSX we just downloaded opens again with the same values
  await page.getByLabel("Open file").setInputFiles(downloaded);
  assert.match(
    await waitForStatus(
      page,
      /^Opened Household budget\.xlsx with \d+ warnings\.$/,
    ),
    /with \d+ warnings/,
  );
  await waitForCell(page, "E2", "2570");
  console.log(
    "PASS Open file reopens an XLSX file with its values and reports its warnings",
  );

  // Open file with CSV
  const csvPath = path.join(workDir, "salaries.csv");
  await writeFile(csvPath, "Item,Amount\nSalary,3000\n");
  await page.getByLabel("Open file").setInputFiles(csvPath);
  assert.equal(
    await waitForStatus(page, /^Opened salaries\.csv/),
    "Opened salaries.csv with 0 warnings.",
  );
  await waitForCell(page, "A2", "Salary");
  assert.equal(await cellText(page, "B2"), "3000");
  console.log("PASS Open file reads a CSV file into a new document");

  // Assistant endpoint: a context from the approved selection, answered by /api/propose
  const context = buildAssistantSelectionContext(workbook, {
    sheetId: workbook.sheets[0].id,
    range: "A1:B5",
  });
  const response = await fetch(`${baseUrl}/api/propose`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ prompt: "Total the amounts", context }),
  });
  assert.equal(response.status, 200);
  const proposal = validateAssistantProposal(await response.json(), context);
  assert.deepEqual(proposal.citations, ["B2", "B3", "B4"]);
  assert.equal(proposal.patches.length, 1);
  assert.equal(proposal.patches[0].address, "B5");
  assert.equal(proposal.patches[0].cell.raw, "=SUM(B2,B3,B4)");
  console.log(
    "PASS /api/propose answers a selection context with a valid proposal that sums B2:B4 into B5",
  );

  // Reset restores the starting values
  await page.getByRole("button", { name: "Reset" }).click();
  assert.equal(
    await waitForStatus(page, /^Reset to the starting budget\.$/),
    "Reset to the starting budget.",
  );
  await waitForCell(page, "E2", "1770");
  assert.equal(await cellText(page, "A2"), "Rent");
  assert.equal(await cellText(page, "B2"), "1200");
  assert.equal(
    await page.getByTestId("workbook-name").textContent(),
    "Household budget",
  );
  console.log("PASS Reset restores the starting values");

  assert.deepEqual(errors, []);
  console.log("PASS no page errors or console errors");
} finally {
  await browser.close();
  await new Promise((resolve) => server.close(resolve));
  await rm(workDir, { recursive: true, force: true });
}
