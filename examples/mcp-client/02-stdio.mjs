/**
 * Example: the reference MCP server as a child process, talked to over stdio.
 *
 * The host writes a small workbook to a folder as native JSON, then starts
 * scripts/mcp-stdio.mjs with --workbook, --sheet, --range and --outbox. Through the
 * SDK's StdioClientTransport it reads the approved range, sees a refusal for a cell
 * outside that range, queues a proposal, and checks that the proposal reached the
 * outbox file. The workbook file is read by the server and never written.
 *
 * Run from the repository root:
 *   node examples/mcp-client/02-stdio.mjs [folder]
 * Without a folder argument, the system temporary folder is used. The working folder
 * created for the run is removed at the end.
 */
import { mkdtemp, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StdioClientTransport } from "@modelcontextprotocol/sdk/client/stdio.js";
import {
  applyCellPatchTransaction,
  createBlankWorkbook,
  writeWorkbookFile,
} from "@ledgerglass/workbook/node";

// The reference server lives in the repository, two folders above this file.
const serverScript = fileURLToPath(
  new URL("../../scripts/mcp-stdio.mjs", import.meta.url),
);

// Same ledger as the in-process example: a header, two amounts, a total and a failing formula.
const blank = createBlankWorkbook("Ledger");
const sheetId = blank.sheets[0].id;
const cells = {
  A1: "Item",
  B1: "Amount",
  A2: "Rent",
  B2: 1200,
  A3: "Power",
  B3: "85.50",
  A4: "Total",
  B4: "=B2+B3",
  C4: "=1/0",
};
const workbook = applyCellPatchTransaction(
  blank,
  Object.entries(cells).map(([address, raw]) => ({
    sheetId,
    address,
    cell: { raw },
  })),
  "Example",
  "Write the ledger",
);

// Work in a new folder so nothing from an earlier run is reused. Removed in `finally`.
const base = resolve(process.argv[2] ?? tmpdir());
const folder = await mkdtemp(join(base, "ledgerglass-stdio-"));
try {
  const workbookPath = join(folder, "workbook.json");
  const outboxPath = join(folder, "proposals.json");
  await writeWorkbookFile(workbook, workbookPath);
  const workbookBefore = await readFile(workbookPath, "utf8");

  // The server is a separate process. Its arguments are the approved selection and
  // the outbox; --sheet takes the stable sheet id from the JSON, not its name.
  const client = new Client({ name: "example-host", version: "1.0.0" });
  await client.connect(
    new StdioClientTransport({
      command: process.execPath,
      args: [
        serverScript,
        "--workbook",
        workbookPath,
        "--sheet",
        sheetId,
        "--range",
        "A1:C4",
        "--outbox",
        outboxPath,
      ],
    }),
  );

  // Calls a tool and returns { ok, data } so a refusal can be shown without throwing.
  async function call(name, args = {}) {
    const result = await client.callTool({ name, arguments: args });
    if (result.isError) return { ok: false, message: result.content[0].text };
    return { ok: true, data: result.structuredContent };
  }

  const { tools } = await client.listTools();
  console.log(
    `tools (${tools.length}): ${tools
      .map((tool) => tool.name)
      .sort()
      .join(", ")}`,
  );

  const inside = await call("workbook_read_range", { range: "A2:B2" });
  console.log(
    `read A2:B2: ${inside.data.cells.map((cell) => cell.raw).join(" | ")}`,
  );

  // D9 is on the same sheet but outside the approved A1:C4, so the server refuses it.
  const outside = await call("workbook_read_range", { range: "D9" });
  console.log(
    `read D9: refused (${outside.ok ? "no" : "yes"}): ${outside.message}`,
  );

  // A proposal must cite the workbook id and revision the model read from the context.
  const context = (await call("workbook_get_context")).data;
  const proposal = {
    id: "example-2",
    workbookId: context.workbookId,
    revision: context.revision,
    explanation: "Store the power amount as a number so the total can use it.",
    citations: ["B3"],
    patches: [{ sheetId, address: "B3", cell: { raw: 85.5 } }],
  };
  const queued = await call("workbook_queue_proposal", { proposal });
  console.log(`queued: ${queued.ok}, applied: ${queued.data.applied}`);

  // The server appends each proposal to the outbox file, wrapped with a timestamp.
  const outbox = JSON.parse(await readFile(outboxPath, "utf8"));
  console.log(`outbox entries: ${outbox.length}`);
  console.log(
    `outbox patches: ${outbox[0].proposal.patches
      .map((patch) => `${patch.address} -> ${patch.cell.raw}`)
      .join(", ")}`,
  );

  // The workbook file is read-only for the server: its bytes are the same as before.
  const workbookAfter = await readFile(workbookPath, "utf8");
  console.log(`workbook file unchanged: ${workbookAfter === workbookBefore}`);

  await client.close();
} finally {
  await rm(folder, { recursive: true, force: true });
}
