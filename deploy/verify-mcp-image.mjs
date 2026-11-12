/**
 * Checks the MCP image from the inside: starts the stdio server on a small workbook and asks it for
 * its tools and context. Run with the image's own Node:
 *   docker run --rm -i --entrypoint node <image> --input-type=module - < deploy/verify-mcp-image.mjs
 */
import assert from "node:assert/strict";
import { mkdtemp, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { pathToFileURL } from "node:url";

const from = (path) => import(pathToFileURL(join(process.cwd(), path)).href);
const { Client } = await from(
  "node_modules/@modelcontextprotocol/sdk/dist/esm/client/index.js",
);
const { StdioClientTransport } = await from(
  "node_modules/@modelcontextprotocol/sdk/dist/esm/client/stdio.js",
);
const { createBlankWorkbook } = await from("src/app/03-workbook-model.js");

const workbook = createBlankWorkbook("Image check");
workbook.sheets[0].cells.A1 = { raw: 2 };
workbook.sheets[0].cells.A2 = { raw: "=A1*21" };
const folder = await mkdtemp(join(tmpdir(), "ledgerglass-mcp-"));
const workbookPath = join(folder, "workbook.json");
await writeFile(workbookPath, JSON.stringify(workbook));

const client = new Client({ name: "image-check", version: "1" });
await client.connect(
  new StdioClientTransport({
    command: process.execPath,
    args: [
      "scripts/mcp-stdio.mjs",
      "--workbook",
      workbookPath,
      "--sheet",
      workbook.sheets[0].id,
      "--range",
      "A1:A2",
      "--outbox",
      join(folder, "proposals.json"),
    ],
  }),
);
try {
  const { tools } = await client.listTools();
  assert.equal(tools.length, 9);
  const context = (
    await client.callTool({ name: "workbook_get_context", arguments: {} })
  ).structuredContent;
  assert.equal(context.cells.find((cell) => cell.address === "A2").value, 42);
  console.log("PASS MCP image serves nine tools and a calculated context");
} finally {
  await client.close();
}
