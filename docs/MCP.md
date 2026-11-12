# MCP server: tools, configuration and integration

Ledgerglass includes a [Model Context Protocol](https://modelcontextprotocol.io) server that lets a model read one approved worksheet selection and queue a proposed change for a person to review. It never changes the workbook.

> The server was rewritten after the original source was lost. It keeps the original entry point (`createWorkbookMcpServer`) and the two tool names that survived (`workbook_get_context`, `workbook_get_capabilities`). The other seven tools were named and specified afresh, so a client written against the original server may need its tool names updated.

## What it can and cannot do

- **Reads:** only cells inside the approved selection (one rectangle on one visible worksheet, at most 500 cells).
- **Writes:** none. A validated proposal is handed to the host's review queue. Applying it is the host's decision.
- **Scope:** a client cannot move or widen the selection. Formulas it evaluates or proposes may only reference cells in the selection and may not reach other worksheets, external workbooks or URLs.

## Running the reference server

```sh
npm run mcp -- \
  --workbook /absolute/path/workbook.json \
  --sheet <worksheet id> \
  --range A1:D20 \
  --outbox /absolute/path/proposals.json
```

| Option       | Meaning                                                                                |
| ------------ | -------------------------------------------------------------------------------------- |
| `--workbook` | Native workbook JSON (File → Save workbook data). Read on each request, never written. |
| `--sheet`    | The worksheet's stable `id` from that JSON, not its display name.                      |
| `--range`    | The approved selection, such as `A1:D20`.                                              |
| `--outbox`   | File that receives validated proposals. Must differ from `--workbook`.                 |

The server speaks MCP over stdio. It needs no build step and no network access. Proposals are appended to the outbox as `{ queuedAt, proposal }` entries.

Example client configuration:

```json
{
  "mcpServers": {
    "ledgerglass": {
      "command": "node",
      "args": [
        "/absolute/path/to/repo/scripts/mcp-stdio.mjs",
        "--workbook",
        "/absolute/path/workbook.json",
        "--sheet",
        "<worksheet id>",
        "--range",
        "A1:D20",
        "--outbox",
        "/absolute/path/proposals.json"
      ]
    }
  }
}
```

The `mcp` Docker target packages the same command:

```sh
docker build --target mcp -t ledgerglass-mcp .
docker run --rm -i -v /absolute/path/data:/data ledgerglass-mcp \
  --workbook /data/workbook.json --sheet <worksheet id> --range A1:D20 --outbox /data/proposals.json
```

## Tools

| Tool                                   | Input                         | Returns                                                                                                                               |
| -------------------------------------- | ----------------------------- | ------------------------------------------------------------------------------------------------------------------------------------- |
| `workbook_get_capabilities`            | none                          | What the server can do, its limits, and that it writes nothing.                                                                       |
| `workbook_get_context`                 | none                          | Workbook id and revision, sheet, range, and each cell's stored content (`raw`) and calculated `value`.                                |
| `workbook_read_range`                  | `range`                       | The cells of a range inside the selection. A range outside it is refused.                                                             |
| `workbook_inspect_formula`             | `formula`                     | Functions and references used, unsupported or volatile functions, and whether the formula could be proposed (with the reason if not). |
| `workbook_evaluate_formula`            | `formula`, optional `address` | The calculated result, without storing the formula.                                                                                   |
| `workbook_get_formula_contracts`       | optional `functions`          | For each function: whether it is supported, its family, whether it is volatile and whether proposals may use it.                      |
| `workbook_check_formula_compatibility` | none                          | Formulas in the selection that use unsupported or volatile functions, reference other sheets or evaluate to an error.                 |
| `workbook_run_audit_checks`            | none                          | Findings that each cite a cell: error values, numbers stored as text, typed-in numbers inside formulas, constants in formula columns. |
| `workbook_queue_proposal`              | `proposal`                    | Confirmation that the proposal was queued. It is not applied.                                                                         |

A failed call returns `isError: true` with a plain-language message, for example "That range is outside the approved selection."

### Proposal format

```json
{
  "id": "p-1",
  "workbookId": "<from workbook_get_context>",
  "revision": 0,
  "explanation": "Why the change is proposed.",
  "citations": ["B3"],
  "patches": [
    { "sheetId": "<sheet id>", "address": "B3", "cell": { "raw": 85.5 } }
  ]
}
```

A proposal is refused if it names another workbook or an older revision, cites or edits a cell outside the selection, edits a cell twice, carries unexpected fields, or contains a formula that uses a volatile or unsupported function or references anything outside the selection. These are the same checks the application applies to its own assistant's proposals.

## Embedding it in your own host

```js
import { createWorkbookMcpServer } from "@ledgerglass/workbook/mcp";

const server = createWorkbookMcpServer({
  getWorkbook: () => authorizedWorkbook, // the workbook this user may read
  getSelection: () => ({ sheetId, range }), // what they approved sharing
  queueProposal: (proposal) => reviewQueue.add(proposal), // shown to a person
});
await server.connect(transport);
```

The host is responsible for authenticating the user, authorizing the workbook and selection, choosing the transport, and reviewing and applying proposals. An optional `beforeRequest` function runs before each tool call, which is where a host can refresh its workbook.

The server module and the model code it imports contain no browser APIs and run in Node 22 or later.
