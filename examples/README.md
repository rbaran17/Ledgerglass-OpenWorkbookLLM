# Examples

Runnable examples of the Ledgerglass SDK. Each imports from the package names a consumer would use (`@ledgerglass/workbook`, `@ledgerglass/workbook/node`, `@ledgerglass/workbook/mcp`); inside this repository those names resolve to the built library, so build it first:

```sh
npm install
npm run build:lib
```

| Folder                              | What it shows                                                                                                                                                        | Run                                                                                 |
| ----------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ----------------------------------------------------------------------------------- |
| [`node-scripts/`](node-scripts)     | Five Node scripts: build and save a workbook, list formula errors in a file, turn a CSV into XLSX, tables and data validation, check an agent's workbook description | `node examples/node-scripts/01-build-and-save.mjs`                                  |
| [`mcp-client/`](mcp-client)         | An MCP client talking to the workbook server in process and over stdio: read the approved range, evaluate a formula, queue a proposal                                | `node examples/mcp-client/01-in-process.mjs`                                        |
| [`react-host/`](react-host)         | A React page that embeds the editor with open, download and reset controls, and a small server that answers the assistant without a model                            | `node examples/react-host/build.mjs && node examples/react-host/server.mjs`         |
| [`source-gateway/`](source-gateway) | A host server with a source gateway (a CSV file and a SQLite database) and a page whose Data > Connections & Refresh dialog previews and loads them                  | `node examples/source-gateway/build.mjs && node examples/source-gateway/server.mjs` |

For an agent that shows a workbook through AG-UI or A2UI, see the host page in [`tests/agent-ui-host/`](../tests/agent-ui-host) and [the agent guide](../docs/AGENT-UI.md).

`npm run test:examples` runs all of them: the Node scripts are compared with their recorded output, and the React host is driven in a browser. The same check runs in CI, so an example that stops working fails the build.
