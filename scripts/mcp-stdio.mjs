#!/usr/bin/env node
/**
 * Reference MCP server over stdio for one workbook file and one approved selection.
 *
 *   npm run mcp -- --workbook /abs/workbook.json --sheet <sheet id> --range A1:D20 --outbox /abs/proposals.json
 *
 * The workbook file is read on every request and never written. Validated proposals are appended to
 * the outbox file for a person to review; nothing is applied.
 */
import { readFile, writeFile, rename } from "node:fs/promises";
import { resolve } from "node:path";
import { parseArgs } from "node:util";

const HELP = `Ledgerglass MCP server (stdio)

Usage:
  node scripts/mcp-stdio.mjs --workbook <file.json> --sheet <sheet id> --range <A1:D20> --outbox <proposals.json>

Options:
  --workbook  Native workbook JSON (File > Save workbook data). Read-only.
  --sheet     The worksheet's stable id from that JSON, not its display name.
  --range     The approved selection, one rectangle such as A1:D20.
  --outbox    File that receives validated proposals for review. Must differ from --workbook.
  --help      Show this text.

The server reads only the approved selection and cannot change the workbook.`;

const fail = (message) => {
  process.stderr.write(`${message}\n\n${HELP}\n`);
  process.exit(2);
};

let options;
try {
  options = parseArgs({
    options: {
      workbook: { type: "string" },
      sheet: { type: "string" },
      range: { type: "string" },
      outbox: { type: "string" },
      help: { type: "boolean", short: "h" },
    },
  }).values;
} catch (error) {
  fail(error.message);
}
if (options.help || process.argv.length <= 2) {
  process.stdout.write(`${HELP}\n`);
  process.exit(0);
}
for (const name of ["workbook", "sheet", "range", "outbox"])
  if (!options[name]) fail(`Missing --${name}.`);
const workbookPath = resolve(options.workbook);
const outboxPath = resolve(options.outbox);
if (workbookPath === outboxPath)
  fail("--workbook and --outbox must be different files.");

const { StdioServerTransport } =
  await import("@modelcontextprotocol/sdk/server/stdio.js");
const { createWorkbookMcpServer } = await import("../src/mcp/server.js");
const { normalizeWorkbookInput } = await import("../src/mcp/model.js");

const getWorkbook = (() => {
  // Re-read only when the file changes, so a calculation is reused between requests.
  let cached = { text: null, workbook: null };
  return async () => {
    const text = await readFile(workbookPath, "utf8");
    if (text !== cached.text)
      cached = { text, workbook: normalizeWorkbookInput(JSON.parse(text)) };
    return cached.workbook;
  };
})();

async function queueProposal(proposal) {
  let queued = [];
  try {
    queued = JSON.parse(await readFile(outboxPath, "utf8"));
    if (!Array.isArray(queued)) throw Error("The outbox file is not a list.");
  } catch (error) {
    if (error.code !== "ENOENT") throw error;
  }
  queued.push({ queuedAt: new Date().toISOString(), proposal });
  const temporary = `${outboxPath}.${process.pid}.tmp`;
  await writeFile(temporary, `${JSON.stringify(queued, null, 2)}\n`, {
    mode: 0o600,
  });
  await rename(temporary, outboxPath);
}

// Fail early with a clear message if the file or selection is unusable.
let current;
try {
  current = await getWorkbook();
} catch (error) {
  fail(`Could not read the workbook: ${error.message}`);
}
const server = createWorkbookMcpServer({
  getWorkbook: () => current,
  getSelection: () => ({ sheetId: options.sheet, range: options.range }),
  queueProposal,
  beforeRequest: async () => {
    current = await getWorkbook();
  },
});
await server.connect(new StdioServerTransport());
