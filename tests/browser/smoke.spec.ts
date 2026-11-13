import { expect, test, type Page } from "@playwright/test";
import { readFile } from "node:fs/promises";

/** Behavioural checks against a built application; they hold for any equivalent build. */

function watchForErrors(page: Page) {
  const errors: string[] = [];
  page.on("pageerror", (error) => errors.push(`pageerror: ${error.message}`));
  page.on("console", (message) => {
    // Browsers request /favicon.ico on their own; the application does not ship one.
    if (message.type() === "error" && !message.location().url.endsWith("/favicon.ico"))
      errors.push(`console: ${message.text()}`);
  });
  return errors;
}

async function open(page: Page, query: string) {
  await page.goto(`/?${query}`);
  await expect(page.getByRole("grid").first()).toBeVisible();
  await expect(page.locator('[data-address="A1"]').first()).toBeVisible();
}

const cell = (page: Page, address: string) =>
  page.locator(`[data-address="${address}"]`).first();

async function select(page: Page, address: string) {
  const box = page.getByLabel("Selected range");
  await box.fill(address);
  await box.press("Enter");
  await expect(cell(page, address.split(":")[0])).toHaveAttribute("aria-selected", "true");
}

test("blank workbook loads under the production content security policy", async ({ page }) => {
  const errors = watchForErrors(page);
  const response = await page.goto("/?temporary");
  expect(response?.headers()["content-security-policy"]).toContain("script-src 'self'");
  await expect(page.getByRole("grid").first()).toBeVisible();
  await expect(page.getByRole("tab", { name: "Home" })).toBeVisible();
  expect(errors).toEqual([]);
});

test("demo workbook calculates cross-sheet formulas", async ({ page }) => {
  const errors = watchForErrors(page);
  await open(page, "demo&temporary");
  await expect(cell(page, "A1")).toHaveAttribute("aria-label", "A1 Order ID");
  await page.getByRole("button", { name: "By region" }).click();
  await expect(cell(page, "A2")).toHaveAttribute("aria-label", "A2 East");
  await expect(cell(page, "B2")).toHaveAttribute("aria-label", "B2 5");
  await expect(cell(page, "D2")).toHaveAttribute("aria-label", "D2 12968.9");
  expect(errors).toEqual([]);
});

test("typed formulas evaluate and recalculate", async ({ page }) => {
  const errors = watchForErrors(page);
  await open(page, "temporary");
  await select(page, "B2");
  await page.keyboard.type("40");
  await page.keyboard.press("Enter");
  await page.keyboard.type("=SUM(B2,2)*2");
  await page.keyboard.press("Enter");
  await expect(cell(page, "B3")).toHaveAttribute("aria-label", "B3 84");
  await select(page, "B2");
  await page.keyboard.type("10");
  await page.keyboard.press("Enter");
  await expect(cell(page, "B3")).toHaveAttribute("aria-label", "B3 24");
  await page.getByRole("button", { name: "Undo" }).click();
  await expect(cell(page, "B3")).toHaveAttribute("aria-label", "B3 84");
  expect(errors).toEqual([]);
});

test("assistant panel loads on demand", async ({ page }) => {
  const errors = watchForErrors(page);
  await open(page, "demo&temporary");
  await page.getByRole("button", { name: "AI assistant" }).click();
  await expect(page.getByLabel("Assistant prompt")).toBeVisible();
  await expect(page.getByRole("button", { name: "New chat" })).toBeVisible();
  expect(errors).toEqual([]);
});

test("workbook health dialog opens", async ({ page }) => {
  const errors = watchForErrors(page);
  await open(page, "demo&temporary");
  await page.getByRole("button", { name: "Workbook health" }).click();
  await expect(page.getByRole("dialog")).toContainText("Diagnostics contain operation categories");
  expect(errors).toEqual([]);
});

test("embedded mode renders the workbook inside the example host", async ({ page }) => {
  const errors = watchForErrors(page);
  await open(page, "embedded&temporary");
  await expect(page.getByRole("tab", { name: "Home" })).toBeVisible();
  expect(errors).toEqual([]);
});

test("CSV export contains the worksheet", async ({ page }) => {
  const errors = watchForErrors(page);
  await open(page, "demo&temporary");
  await page.getByRole("button", { name: "File", exact: true }).click();
  const downloading = page.waitForEvent("download");
  await page.getByText("Save as CSV").click();
  await page.getByRole("button", { name: "Download reviewed data" }).click();
  const download = await downloading;
  const text = await readFile(await download.path(), "utf8");
  expect(text).toContain("Order ID");
  expect(errors).toEqual([]);
});

test("XLSX export can be opened again", async ({ page }) => {
  const errors = watchForErrors(page);
  const workers: string[] = [];
  page.on("worker", (w) => workers.push(w.url()));
  await open(page, "demo&temporary");
  await page.getByRole("button", { name: "File", exact: true }).click();
  const downloading = page.waitForEvent("download");
  await page.getByText("Save as Excel").click();
  await page.getByRole("button", { name: "Download derived XLSX" }).click();
  const download = await downloading;
  const bytes = await readFile(await download.path());
  expect(bytes.subarray(0, 2).toString()).toBe("PK");

  await open(page, "temporary");
  await page.getByLabel("Import workbook file").setInputFiles({
    name: "roundtrip.xlsx",
    mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    buffer: bytes,
  });
  await page.getByRole("button", { name: "Open derived workbook" }).click();
  await expect(cell(page, "A1")).toHaveAttribute("aria-label", "A1 Order ID", { timeout: 30_000 });
  expect(workers.length).toBeGreaterThan(0);
  expect(errors).toEqual([]);
});

test("CSV import keeps quoted fields and numbers", async ({ page }) => {
  const errors = watchForErrors(page);
  await open(page, "temporary");
  await page.getByLabel("Import workbook file").setInputFiles({
    name: "ledger.csv",
    mimeType: "text/csv",
    buffer: Buffer.from('Account,Amount\n"Cash, operating",1250.5\nPayables,-300\n'),
  });
  await page.getByRole("button", { name: "Import reviewed data" }).click();
  await expect(cell(page, "A2")).toHaveAttribute("aria-label", "A2 Cash, operating");
  await expect(cell(page, "B2")).toHaveAttribute("aria-label", "B2 1250.5");
  await expect(cell(page, "B3")).toHaveAttribute("aria-label", "B3 -300");
  expect(errors).toEqual([]);
});

test("native workbook data round-trips and invalid data is refused", async ({ page }) => {
  const errors = watchForErrors(page);
  await open(page, "demo&temporary");
  await page.getByRole("button", { name: "File", exact: true }).click();
  const downloading = page.waitForEvent("download");
  await page.getByText("Save workbook data").click();
  const saved = await readFile(await (await downloading).path(), "utf8");
  expect(JSON.parse(saved).schemaVersion).toBe(1);

  await open(page, "temporary");
  await page.getByLabel("Import workbook file").setInputFiles({
    name: "workbook.json",
    mimeType: "application/json",
    buffer: Buffer.from(saved),
  });
  await expect(cell(page, "A1")).toHaveAttribute("aria-label", "A1 Order ID");

  await open(page, "temporary");
  await page.getByLabel("Import workbook file").setInputFiles({
    name: "invalid.json",
    mimeType: "application/json",
    buffer: Buffer.from(saved.replace('"sheets"', '"sheetz"')),
  });
  await expect(page.getByText("A workbook must contain at least one sheet.")).toBeVisible();
  await expect(cell(page, "A1")).toHaveAttribute("aria-label", /^A1\s*$/);
  expect(errors).toEqual([]);
});

test("XML import parses elements and entities", async ({ page }) => {
  const errors = watchForErrors(page);
  await open(page, "temporary");
  await page.getByRole("button", { name: "File", exact: true }).click();
  await page.getByText(/^Import sources/).click();
  await page.getByLabel("Import source file").setInputFiles({
    name: "ledger.xml",
    mimeType: "application/xml",
    buffer: Buffer.from(
      '<?xml version="1.0"?><rows><row><Account>Cash &amp; bank</Account><Amount>1250.5</Amount></row><row><Account>Payables</Account><Amount>-300</Amount></row></rows>',
    ),
  });
  await expect(page.getByText("1 table candidates")).toBeVisible();
  await page.getByRole("button", { name: "Preview import", exact: true }).click();
  await page.getByRole("button", { name: "Import reviewed table", exact: true }).click();
  await expect(cell(page, "A2")).toHaveAttribute("aria-label", "A2 Cash & bank");
  await expect(cell(page, "B2")).toHaveAttribute("aria-label", "B2 1250.5");
  await expect(cell(page, "B3")).toHaveAttribute("aria-label", "B3 -300");
  expect(errors).toEqual([]);
});

test("alternate workbook format round-trips through its ZIP archive", async ({ page }) => {
  const errors = watchForErrors(page);
  await open(page, "demo&temporary");
  await page.getByRole("button", { name: "File", exact: true }).click();
  await page.getByText(/^Workbook formats/).click();
  await page.getByRole("button", { name: "Review XLTX export" }).click();
  const downloading = page.waitForEvent("download");
  await page.getByRole("button", { name: "Download derived XLTX" }).click();
  const download = await downloading;
  const bytes = await readFile(await download.path());
  expect(bytes.subarray(0, 2).toString()).toBe("PK");

  await open(page, "temporary");
  await page.getByLabel("Import workbook file").setInputFiles({
    name: "roundtrip.xltx",
    mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.template",
    buffer: bytes,
  });
  await page.getByRole("button", { name: "Preview template workbook" }).click();
  await page.getByRole("button", { name: "Open derived workbook" }).click();
  await expect(cell(page, "A1")).toHaveAttribute("aria-label", "A1 Order ID", { timeout: 30_000 });
  expect(errors).toEqual([]);
});

test("schedule PDF export produces a PDF file", async ({ page }) => {
  const errors = watchForErrors(page);
  await open(page, "demo&temporary");
  await page.getByRole("button", { name: "By region" }).click();
  // Column widths depend on the fonts installed; the app refuses to print a number
  // that does not fit, so size the columns to their content first.
  await select(page, "A1:E6");
  await page.getByRole("button", { name: "AutoFit Column Width" }).click();
  await page.getByRole("tab", { name: "Page Layout", exact: true }).click();
  await page.getByRole("button", { name: "Margins", exact: true }).click();
  const dialog = page.getByRole("dialog", { name: "Page setup and print preview" });
  await dialog.getByRole("button", { name: "Create PDF", exact: true }).click();
  const link = dialog.getByRole("link", { name: /Download PDF/ });
  // If no PDF appears, say what the dialog reported instead of only timing out.
  const appeared = await link
    .waitFor({ state: "visible", timeout: 45_000 })
    .then(() => true, () => false);
  if (!appeared) {
    const report = (await dialog.innerText())
      .split("\n")
      .filter((line) => /issue|warning|error|fail|block|cannot|unable|exceed|beyond/i.test(line))
      .slice(0, 12)
      .join(" | ");
    throw new Error(`No PDF was produced. Dialog reported: ${report || "(nothing matching)"}; page errors: ${errors.join("; ") || "none"}`);
  }
  const downloading = page.waitForEvent("download");
  await link.click();
  const download = await downloading;
  const bytes = await readFile(await download.path());
  expect(bytes.subarray(0, 5).toString()).toBe("%PDF-");
  expect(bytes.length).toBeGreaterThan(1000);
  expect(bytes.subarray(-1024).toString("latin1")).toContain("%%EOF");
  expect(errors).toEqual([]);
});

test("assistant reply is rendered as Markdown", async ({ page }) => {
  const errors = watchForErrors(page);
  await open(page, "demo&temporary");
  await page.getByRole("button", { name: "AI assistant" }).click();
  const prompt = page.getByLabel("Assistant prompt");
  await prompt.fill("Summarize the selected range");
  await prompt.press("Enter");
  const conversation = page.getByLabel("Conversation");
  await expect(conversation).toContainText("Local deterministic demo", { timeout: 30_000 });
  await expect(conversation.locator("strong").first()).toBeVisible();
  await expect(conversation.locator("p").first()).toBeVisible();
  expect(errors).toEqual([]);
});

test("assistant keeps a multi-turn conversation", async ({ page }) => {
  const errors = watchForErrors(page);
  await open(page, "demo&temporary");
  await page.getByRole("button", { name: "AI assistant" }).click();
  const prompt = page.getByLabel("Assistant prompt");
  const conversation = page.getByLabel("Conversation");
  await prompt.fill("First question about totals");
  await prompt.press("Enter");
  await expect(conversation.getByText(/Local deterministic demo/)).toHaveCount(1, { timeout: 30_000 });
  await prompt.fill("Second question about accruals");
  await prompt.press("Enter");
  await expect(conversation.getByText(/Local deterministic demo/)).toHaveCount(2, { timeout: 30_000 });
  await expect(conversation).toContainText("First question about totals");
  await expect(conversation).toContainText("Second question about accruals");
  await expect(conversation.getByLabel("Your message", { exact: true })).toHaveCount(2);
  await expect(conversation.getByLabel("Assistant message", { exact: true })).toHaveCount(2);
  expect(errors).toEqual([]);
});

test("assistant starts a new chat", async ({ page }) => {
  const errors = watchForErrors(page);
  await open(page, "demo&temporary");
  await page.getByRole("button", { name: "AI assistant" }).click();
  const conversation = page.getByLabel("Conversation");
  await page.getByLabel("Assistant prompt").fill("Question before the reset");
  await page.getByLabel("Assistant prompt").press("Enter");
  await expect(conversation.getByText(/Local deterministic demo/)).toHaveCount(1, { timeout: 30_000 });
  await page.getByRole("button", { name: "New chat" }).click();
  await expect(conversation).not.toContainText("Question before the reset");
  await expect(conversation.getByLabel("Assistant message", { exact: true })).toHaveCount(0);
  await expect(conversation).toContainText("How can I help?");
  expect(errors).toEqual([]);
});

test("assistant accepts a file attachment", async ({ page }) => {
  const errors = watchForErrors(page);
  await open(page, "demo&temporary");
  await page.getByRole("button", { name: "AI assistant" }).click();
  const conversation = page.getByLabel("Conversation");
  await page.getByLabel("Upload chat attachments").setInputFiles({
    name: "notes.txt",
    mimeType: "text/plain",
    buffer: Buffer.from("Quarter-end accrual notes"),
  });
  await expect(page.getByRole("button", { name: "Remove attachment notes.txt" })).toBeVisible();
  await expect(page.getByText("notes.txt", { exact: true })).toBeVisible();
  await page.getByLabel("Assistant prompt").fill("Summarize the attachment");
  await page.getByLabel("Assistant prompt").press("Enter");
  await expect(conversation.getByText(/Local deterministic demo/)).toHaveCount(1, { timeout: 30_000 });
  await expect(conversation).toContainText("Summarize the attachment");
  await expect(conversation).toContainText("notes.txt");
  expect(errors).toEqual([]);
});

test("assistant panel can be closed and reopened", async ({ page }) => {
  const errors = watchForErrors(page);
  await open(page, "demo&temporary");
  await page.getByRole("button", { name: "AI assistant" }).click();
  const prompt = page.getByLabel("Assistant prompt");
  const conversation = page.getByLabel("Conversation");
  await prompt.fill("Question before closing");
  await prompt.press("Enter");
  await expect(conversation.getByText(/Local deterministic demo/)).toHaveCount(1, { timeout: 30_000 });
  await page.getByRole("button", { name: "Close side panel" }).click();
  await expect(prompt).toBeHidden();
  await page.getByRole("button", { name: "AI assistant" }).click();
  await expect(prompt).toBeVisible();
  // Observed: the conversation survives closing the panel within the same page session.
  await expect(conversation).toContainText("Question before closing");
  await expect(conversation.getByText(/Local deterministic demo/)).toHaveCount(1);
  expect(errors).toEqual([]);
});
