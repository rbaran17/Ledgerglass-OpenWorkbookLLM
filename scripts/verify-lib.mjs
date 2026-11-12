/**
 * Checks the built library the way a host application would use it: bundles a small React page
 * that imports lib/ledgerglass.js, serves it with the production runtime (and its content security
 * policy), and drives it in a browser. Run `npm run build:lib` first.
 */
import assert from "node:assert/strict";
import { mkdir, rm, writeFile } from "node:fs/promises";
import { build } from "esbuild";
import { chromium } from "playwright";
import { startWebServer } from "../deploy/server.mjs";

const hostDirectory = "work/lib-host";
await rm(hostDirectory, { recursive: true, force: true });
await mkdir(hostDirectory, { recursive: true });
await writeFile(
  `${hostDirectory}/host.jsx`,
  `import { createRoot } from "react-dom/client";
import { AuditWorkbook, createBlankWorkbook, createHttpAdapter, localDemoAdapter, validateAssistantProposal } from "../../lib/ledgerglass.js";
import "../../lib/ledgerglass.css";

const workbook = createBlankWorkbook("Host workbook");
workbook.sheets[0].cells.A1 = { raw: "From the host" };
window.hostChanges = 0;
window.libraryExports = [typeof createHttpAdapter, typeof localDemoAdapter.propose, typeof validateAssistantProposal];
createRoot(document.getElementById("host")).render(
  <div style={{ height: "600px" }}>
    <AuditWorkbook initialWorkbook={workbook} adapter={localDemoAdapter} onChange={() => { window.hostChanges += 1; }} />
  </div>,
);
`,
);
await build({
  entryPoints: [`${hostDirectory}/host.jsx`],
  outdir: `${hostDirectory}/public/assets`,
  bundle: true,
  format: "esm",
  splitting: true,
  jsx: "automatic",
  minify: true,
  define: { "process.env.NODE_ENV": '"production"' },
  logLevel: "error",
});
await writeFile(
  `${hostDirectory}/public/index.html`,
  `<!doctype html><html lang="en"><head><meta charset="UTF-8"/><title>Library host</title><link rel="stylesheet" href="assets/host.css"/></head><body><div id="host"></div><script type="module" src="assets/host.js"></script></body></html>`,
);

const runtime = await startWebServer({
  root: `${hostDirectory}/public`,
  host: "127.0.0.1",
  port: 0,
  logger: () => {},
});
const browser = await chromium.launch(
  process.env.CHROME_PATH ? { executablePath: process.env.CHROME_PATH } : {},
);
try {
  const page = await browser.newPage();
  const errors = [];
  page.on("pageerror", (error) => errors.push(error.message));
  await page.goto(`http://127.0.0.1:${runtime.address.port}/`);
  const cell = (address) => page.locator(`[data-address="${address}"]`).first();
  await cell("A1").waitFor();
  assert.match(await cell("A1").getAttribute("aria-label"), /From the host/);
  console.log(
    "PASS the editor mounts from the library with the host's workbook",
  );

  await cell("B2").dblclick();
  await page.keyboard.type("=6*7");
  await page.keyboard.press("Enter");
  await page.waitForFunction(() =>
    document
      .querySelector('[data-address="B2"]')
      ?.getAttribute("aria-label")
      ?.includes("42"),
  );
  assert.ok((await page.evaluate(() => window.hostChanges)) > 0);
  console.log("PASS formulas calculate and the host is told about changes");

  assert.deepEqual(await page.evaluate(() => window.libraryExports), [
    "function",
    "function",
    "function",
  ]);
  assert.deepEqual(errors, []);
  console.log(
    "PASS adapters and protocol functions are exported; no page errors",
  );
} finally {
  await browser.close();
  await runtime.close();
  await rm(hostDirectory, { recursive: true, force: true });
}
