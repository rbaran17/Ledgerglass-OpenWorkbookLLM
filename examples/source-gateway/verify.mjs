/**
 * Drives the editor's own Data > Connections & Refresh dialog against this example's gateway:
 * the recovered client on one side, the new server on the other.
 *
 *   node examples/source-gateway/verify.mjs      (after `npm run build:lib`)
 */
import assert from "node:assert/strict";
import { cp, mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import { chromium } from "playwright";
import { buildPage } from "./build.mjs";
import { startServer } from "./server.mjs";

const here = fileURLToPath(new URL(".", import.meta.url));
await buildPage();
// Work on a copy of the data so the check can change a source file.
const dataDirectory = await mkdtemp(join(tmpdir(), "ledgerglass-sources-"));
await cp(join(here, "data/invoices.csv"), join(dataDirectory, "invoices.csv"));
const server = await startServer({ port: 0, dataDirectory });
const origin = `http://127.0.0.1:${server.address().port}`;
const browser = await chromium.launch(
  process.env.CHROME_PATH ? { executablePath: process.env.CHROME_PATH } : {},
);

try {
  const page = await browser.newPage({
    viewport: { width: 1400, height: 950 },
  });
  const pageErrors = [];
  page.on("pageerror", (error) => pageErrors.push(error.message));
  const requests = [];
  page.on("request", (request) => {
    if (request.url().endsWith("/api/sources"))
      requests.push(JSON.parse(request.postData()));
  });
  const cellText = async (address) =>
    (await page
      .locator(`[data-address="${address}"]`)
      .first()
      .getAttribute("aria-label")) ?? "";
  const expectCell = async (address, pattern) => {
    await page.waitForFunction(
      ([selector, source]) =>
        new RegExp(source).test(
          document.querySelector(selector)?.getAttribute("aria-label") ?? "",
        ),
      [`[data-address="${address}"]`, pattern.source],
      { timeout: 10000 },
    );
  };
  async function openConnections() {
    // After refreshing a saved connection the dialog stays open with the result.
    const open = page.getByRole("dialog").last();
    if (await open.isVisible().catch(() => false)) return open;
    await page.getByRole("tab", { name: "Data", exact: true }).click();
    await page
      .locator(".xl-ribbon-body button", { hasText: "Connections & Refresh" })
      .first()
      .click();
    const dialog = page.getByRole("dialog").last();
    await dialog.getByLabel("Approved source", { exact: true }).waitFor();
    return dialog;
  }
  const dialogText = async (dialog) =>
    (await dialog.innerText()).replace(/\s+/g, " ");

  await page.goto(origin);
  await page.locator('[data-address="A1"]').first().waitFor();

  // 1. The catalog.
  let dialog = await openConnections();
  const sourceOptions = await dialog
    .getByLabel("Approved source", { exact: true })
    .locator("option")
    .allTextContents();
  assert.deepEqual(sourceOptions, [
    "Choose a source",
    "Open invoices (CSV) · file-folder",
    "General ledger (SQLite) · database",
  ]);
  assert.deepEqual(requests, [{ operation: "catalog" }]);
  console.log("PASS the dialog lists the gateway's two sources");

  // 2. Preview and apply the CSV source.
  await dialog
    .getByLabel("Approved source", { exact: true })
    .selectOption({ index: 1 });
  await dialog.getByRole("button", { name: "Preview refresh" }).click();
  await dialog.getByText("Refresh preview · 4 records").waitFor();
  let text = await dialogText(dialog);
  assert.match(text, /Rows: 0 → 4/);
  assert.match(text, /Amount 0 4495\.75 4495\.75 4 \/ 0/);
  assert.match(text, /INV-1003 2026-01-14 Fabrikam 845\.25 true/);
  console.log(
    "PASS the preview shows the rows and the control total (4495.75)",
  );

  await dialog.getByRole("button", { name: "Apply reviewed refresh" }).click();
  await expectCell("A2", /INV-1001/);
  assert.match(await cellText("A1"), /Invoice/);
  assert.match(await cellText("B2"), /2026-01-05/);
  assert.match(await cellText("D2"), /1200\.5/);
  assert.match(await cellText("D5"), /2150/);
  assert.deepEqual(
    requests.slice(1).map((request) => request.operation),
    ["authorize", "refresh", "authorize"],
    "preview asks for an approval and the data; apply asks for a fresh approval",
  );
  assert.ok(
    requests.slice(1).every((request) => request.sourceId === "invoices"),
  );
  console.log(
    "PASS applying places the rows in a worksheet; apply checks access again",
  );

  // 3. The source changes; a refresh shows the difference before anything is replaced.
  await writeFile(
    join(dataDirectory, "invoices.csv"),
    "Invoice,Date,Customer,Amount,Paid\nINV-1001,2026-01-05,Northwind,1200.50,true\nINV-1005,2026-02-02,Contoso,99.50,false\n",
  );
  dialog = await openConnections();
  const connections = await dialog
    .getByLabel("Connection", { exact: true })
    .locator("option")
    .allTextContents();
  assert.equal(
    connections.length,
    2,
    "the applied source is now a saved connection",
  );
  await dialog
    .getByLabel("Connection", { exact: true })
    .selectOption({ index: 1 });
  await dialog.getByRole("button", { name: "Preview refresh" }).click();
  await dialog.getByText("Refresh preview · 2 records").waitFor();
  text = await dialogText(dialog);
  assert.match(text, /Rows: 4 → 2/);
  assert.match(text, /Amount 4495\.75 1300 -3195\.75/);
  await dialog.getByRole("button", { name: "Apply reviewed refresh" }).click();
  await expectCell("A3", /INV-1005/);
  console.log(
    "PASS a changed source previews as 4 → 2 rows with the control total change, then refreshes",
  );

  // 4. A source whose columns no longer match what was approved is refused, and the data is kept.
  await writeFile(
    join(dataDirectory, "invoices.csv"),
    "Invoice,Date,Customer,Amount,Paid\nINV-1001,2026-01-05,Northwind,not a number,true\n",
  );
  dialog = await openConnections();
  await dialog
    .getByLabel("Connection", { exact: true })
    .selectOption({ index: 1 });
  await dialog.getByRole("button", { name: "Preview refresh" }).click();
  await dialog
    .getByText(/source schema changed/i)
    .first()
    .waitFor();
  assert.match(
    await dialogText(dialog),
    /Stale · previous successful data retained/,
  );
  assert.equal(
    await dialog
      .getByRole("button", { name: "Apply reviewed refresh" })
      .isEnabled(),
    false,
  );
  await dialog.getByRole("button", { name: "Close" }).first().click();
  assert.match(await cellText("A3"), /INV-1005/);
  console.log(
    "PASS a value of the wrong type is refused as a schema change and the worksheet is kept",
  );

  assert.deepEqual(pageErrors, []);
  console.log("PASS no page errors");
} finally {
  await browser.close();
  server.close();
  await rm(dataDirectory, { recursive: true, force: true });
}
