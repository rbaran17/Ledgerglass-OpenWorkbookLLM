# MCP client examples

Two runnable Node scripts that drive the Ledgerglass MCP server from a host, the way an application or an agent runtime would. Both build the same small ledger in code, approve the selection A1:C4, and call the nine read-only tools, then queue one proposal for review. The first script runs the server in the same process and connects with the SDK's in-memory transport. The second starts the reference server, `scripts/mcp-stdio.mjs`, as a child process and talks to it over stdio. Each script has a `.expected.txt` file with its exact stdout. Run the scripts from the repository root, with Node 22 or later and the repository's dependencies installed.

## 01-in-process.mjs

Creates the server with `createWorkbookMcpServer` from `@ledgerglass/workbook/mcp`, connects an MCP `Client` through `InMemoryTransport`, and lists the tools. It then gets the capabilities and context, reads A2:B2, inspects and evaluates a formula, runs the audit checks (which flag a number stored as text and a failing formula), and queues a proposal. The proposal reaches the host's review queue, and the workbook is unchanged.

```sh
node examples/mcp-client/01-in-process.mjs
```

```text
tools (9):
  workbook_check_formula_compatibility
  workbook_evaluate_formula
  workbook_get_capabilities
  workbook_get_context
  workbook_get_formula_contracts
  workbook_inspect_formula
  workbook_queue_proposal
  workbook_read_range
  workbook_run_audit_checks
max selection cells: 500
writes: None. Proposals are queued for human review and are never applied by this server.
context: sheet Sheet1, range A1:C4, 12 cells, revision 1
read A2:B2: Rent | 1200
inspect =SUM(B2:B3): functions SUM, references B2, B3, proposable true
evaluate =B2*2 as B4: 2400
audit: 12 cells checked, 2 findings
  B3 number-stored-as-text (medium)
  C4 error-value (high)
queued: true, applied: false
host review queue: 1 proposal
queued patches: B3 -> 85.5
workbook unchanged: true
B3 still holds: 85.50
```

## 02-stdio.mjs

Writes the ledger to a working folder as native JSON with `writeWorkbookFile`, then starts `scripts/mcp-stdio.mjs` with `--workbook`, `--sheet`, `--range A1:C4` and `--outbox`, connected through `StdioClientTransport`. It reads the approved range, shows the refusal for D9 (outside the selection), queues a proposal, and checks the outbox file. The working folder is removed at the end. An optional folder argument chooses where the working folder is created.

```sh
node examples/mcp-client/02-stdio.mjs
```

```text
tools (9): workbook_check_formula_compatibility, workbook_evaluate_formula, workbook_get_capabilities, workbook_get_context, workbook_get_formula_contracts, workbook_inspect_formula, workbook_queue_proposal, workbook_read_range, workbook_run_audit_checks
read A2:B2: Rent | 1200
read D9: refused (yes): That range is outside the approved selection.
queued: true, applied: false
outbox entries: 1
outbox patches: B3 -> 85.5
workbook file unchanged: true
```
