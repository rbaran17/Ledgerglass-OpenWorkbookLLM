import { afterEach, describe, expect, test } from "vitest";
import { mkdtemp, readFile, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import { StdioClientTransport } from "@modelcontextprotocol/sdk/client/stdio.js";
import { createBlankWorkbook } from "../../src/app/03-workbook-model.js";
import { createWorkbookMcpServer } from "../../src/mcp/server.js";

function ledger() {
  const workbook = createBlankWorkbook("Ledger");
  const sheet = workbook.sheets[0];
  Object.assign(sheet.cells, {
    A1: { raw: "Item" },
    B1: { raw: "Amount" },
    A2: { raw: "Rent" },
    B2: { raw: 1200 },
    A3: { raw: "Power" },
    B3: { raw: "85.50" },
    A4: { raw: "Total" },
    B4: { raw: "=B2+B3" },
    C4: { raw: "=1/0" },
    D9: { raw: "outside the selection" },
  });
  return workbook;
}

const open = [];
afterEach(async () => {
  for (const close of open.splice(0)) await close();
});

async function connect(workbook, range = "A1:C4") {
  const queued = [];
  const server = createWorkbookMcpServer({
    getWorkbook: () => workbook,
    getSelection: () => ({ sheetId: workbook.sheets[0].id, range }),
    queueProposal: (proposal) => queued.push(proposal),
  });
  const client = new Client({ name: "unit", version: "1" });
  const [a, b] = InMemoryTransport.createLinkedPair();
  await server.connect(a);
  await client.connect(b);
  open.push(async () => {
    await client.close();
    await server.close();
  });
  const call = async (name, args = {}) => {
    const result = await client.callTool({ name, arguments: args });
    return result.isError
      ? { error: result.content[0].text }
      : result.structuredContent;
  };
  return { client, call, queued };
}

const proposalFor = (workbook, patches, extra = {}) => ({
  id: "p-1",
  workbookId: workbook.id,
  revision: workbook.revision,
  explanation: "Store the amount as a number.",
  citations: ["B3"],
  patches,
  ...extra,
});

describe("MCP server", () => {
  test("offers nine tools", async () => {
    const { client } = await connect(ledger());
    const { tools } = await client.listTools();
    expect(tools.map((tool) => tool.name).sort()).toEqual([
      "workbook_check_formula_compatibility",
      "workbook_evaluate_formula",
      "workbook_get_capabilities",
      "workbook_get_context",
      "workbook_get_formula_contracts",
      "workbook_inspect_formula",
      "workbook_queue_proposal",
      "workbook_read_range",
      "workbook_run_audit_checks",
    ]);
  });

  test("capabilities state the limits and that nothing is written", async () => {
    const { call } = await connect(ledger());
    const capabilities = await call("workbook_get_capabilities");
    expect(capabilities.limits.maxSelectionCells).toBe(500);
    expect(capabilities.authority.writes).toMatch(/never applied/);
    expect(capabilities.customFunctions.execution).toMatch(/cannot install/);
  });

  test("context holds only the approved selection, with calculated values", async () => {
    const workbook = ledger();
    const { call } = await connect(workbook);
    const context = await call("workbook_get_context");
    expect(context).toMatchObject({
      workbookId: workbook.id,
      revision: workbook.revision,
      sheetName: "Sheet1",
      range: "A1:C4",
    });
    expect(context.cells).toHaveLength(12);
    const total = context.cells.find((cell) => cell.address === "B4");
    expect(total.raw).toBe("=B2+B3");
    expect(JSON.stringify(context)).not.toContain("outside the selection");
  });

  test("reads a range inside the selection and refuses one outside it", async () => {
    const { call } = await connect(ledger());
    const inside = await call("workbook_read_range", { range: "A2:B2" });
    expect(inside.cells.map((cell) => cell.raw)).toEqual(["Rent", 1200]);
    const outside = await call("workbook_read_range", { range: "D9" });
    expect(outside.error).toMatch(/outside the approved selection/);
  });

  test("inspects a formula and says whether it could be proposed", async () => {
    const { call } = await connect(ledger());
    const good = await call("workbook_inspect_formula", {
      formula: "=SUM(B2:B3)",
    });
    expect(good).toMatchObject({
      functions: ["SUM"],
      references: ["B2", "B3"],
      proposable: true,
    });
    const outside = await call("workbook_inspect_formula", {
      formula: "=D9*2",
    });
    expect(outside.proposable).toBe(false);
    expect(outside.reason).toMatch(/approved selection/);
    const volatile = await call("workbook_inspect_formula", {
      formula: "=NOW()",
    });
    expect(volatile.volatileFunctions).toEqual(["NOW"]);
    expect(volatile.proposable).toBe(false);
  });

  test("evaluates a formula without storing it", async () => {
    const workbook = ledger();
    const before = JSON.stringify(workbook);
    const { call } = await connect(workbook);
    const result = await call("workbook_evaluate_formula", {
      formula: "=B2*2",
    });
    expect(result.result).toBe(2400);
    expect(JSON.stringify(workbook)).toBe(before);
    const refused = await call("workbook_evaluate_formula", {
      formula: "=D9",
    });
    expect(refused.error).toMatch(/approved selection/);
  });

  test("describes supported functions", async () => {
    const { call } = await connect(ledger());
    const { functions } = await call("workbook_get_formula_contracts", {
      functions: ["pmt", "NOW", "NOT_A_FUNCTION"],
    });
    expect(functions).toEqual([
      {
        name: "PMT",
        supported: true,
        family: "financial",
        volatile: false,
        allowedInProposals: true,
      },
      expect.objectContaining({
        name: "NOW",
        supported: true,
        volatile: true,
        allowedInProposals: false,
      }),
      expect.objectContaining({ name: "NOT_A_FUNCTION", supported: false }),
    ]);
  });

  test("reports formulas that evaluate to an error", async () => {
    const { call } = await connect(ledger());
    const report = await call("workbook_check_formula_compatibility");
    expect(report.formulaCount).toBe(2);
    expect(report.issues).toEqual([
      {
        address: "C4",
        formula: "=1/0",
        problems: ["Evaluates to #DIV/0!."],
      },
    ]);
  });

  test("audit checks cite the cells they are about", async () => {
    const { call } = await connect(ledger());
    const { findings } = await call("workbook_run_audit_checks");
    expect(findings).toContainEqual(
      expect.objectContaining({ check: "error-value", address: "C4" }),
    );
    expect(findings).toContainEqual(
      expect.objectContaining({
        check: "number-stored-as-text",
        address: "B3",
      }),
    );
  });

  test("queues a valid proposal and leaves the workbook unchanged", async () => {
    const workbook = ledger();
    const before = JSON.stringify(workbook);
    const { call, queued } = await connect(workbook);
    const sheetId = workbook.sheets[0].id;
    const result = await call("workbook_queue_proposal", {
      proposal: proposalFor(workbook, [
        { sheetId, address: "B3", cell: { raw: 85.5 } },
      ]),
    });
    expect(result).toMatchObject({ queued: true, applied: false, id: "p-1" });
    expect(queued).toHaveLength(1);
    expect(queued[0].patches[0]).toMatchObject({
      address: "B3",
      cell: { raw: 85.5 },
    });
    expect(JSON.stringify(workbook)).toBe(before);
  });

  test("refuses proposals outside the selection, stale ones and unsafe formulas", async () => {
    const workbook = ledger();
    const { call, queued } = await connect(workbook);
    const sheetId = workbook.sheets[0].id;
    const outside = await call("workbook_queue_proposal", {
      proposal: proposalFor(workbook, [
        { sheetId, address: "D9", cell: { raw: 1 } },
      ]),
    });
    expect(outside.error).toMatch(/inside the approved selection/);
    const stale = await call("workbook_queue_proposal", {
      proposal: proposalFor(
        workbook,
        [{ sheetId, address: "B3", cell: { raw: 1 } }],
        { revision: workbook.revision + 1 },
      ),
    });
    expect(stale.error).toMatch(/stale/);
    const unsafe = await call("workbook_queue_proposal", {
      proposal: proposalFor(workbook, [
        { sheetId, address: "B3", cell: { raw: "=Sheet2!A1" } },
      ]),
    });
    expect(unsafe.error).toBeTruthy();
    expect(queued).toHaveLength(0);
  });

  test("needs the three host functions", () => {
    expect(() => createWorkbookMcpServer({ getWorkbook: () => ({}) })).toThrow(
      /getSelection/,
    );
  });
});

describe("MCP stdio command", () => {
  test("serves a workbook file and appends proposals to the outbox", async () => {
    const workbook = ledger();
    const folder = await mkdtemp(join(tmpdir(), "ledgerglass-mcp-"));
    const workbookPath = join(folder, "workbook.json");
    const outboxPath = join(folder, "proposals.json");
    await writeFile(workbookPath, JSON.stringify(workbook));
    const client = new Client({ name: "unit-stdio", version: "1" });
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
          "A1:C4",
          "--outbox",
          outboxPath,
        ],
      }),
    );
    open.push(() => client.close());
    const context = (
      await client.callTool({ name: "workbook_get_context", arguments: {} })
    ).structuredContent;
    expect(context.cells).toHaveLength(12);
    const queued = await client.callTool({
      name: "workbook_queue_proposal",
      arguments: {
        proposal: proposalFor(workbook, [
          {
            sheetId: workbook.sheets[0].id,
            address: "B3",
            cell: { raw: 85.5 },
          },
        ]),
      },
    });
    expect(queued.structuredContent.queued).toBe(true);
    const outbox = JSON.parse(await readFile(outboxPath, "utf8"));
    expect(outbox).toHaveLength(1);
    expect(outbox[0].proposal.id).toBe("p-1");
    expect(JSON.parse(await readFile(workbookPath, "utf8"))).toEqual(
      JSON.parse(JSON.stringify(workbook)),
    );
  });
});
