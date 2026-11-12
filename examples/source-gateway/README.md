# Source gateway example

A host server that offers two data sources to the editor, and a page that uses them. The editor's **Data > Connections & Refresh** dialog lists the sources, previews one with its control totals, places it in a worksheet and refreshes it later. The browser only names a source; the server decides who may see it and how it is read. See [the sources guide](../../docs/SOURCES.md).

| File                | What it is                                                                                                                                                                            |
| ------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `server.mjs`        | The host: a source gateway over a CSV file and a SQLite database, plus the static page. `authenticate` and `isAllowed` are where a real host checks its session and its access rules. |
| `src/main.jsx`      | The page: `AuditWorkbook` with `createHttpSourceAdapter({ endpoint: "/api/sources" })` as its `externalSourceAdapter`.                                                                |
| `data/invoices.csv` | The CSV source. The SQLite database is created at start-up.                                                                                                                           |
| `build.mjs`         | Bundles the page into `public/` (generated, not committed).                                                                                                                           |
| `verify.mjs`        | Drives the dialog in a browser and checks each step.                                                                                                                                  |

## Run it

From the repository root, after `npm install` and `npm run build:lib`:

```sh
node examples/source-gateway/build.mjs
node examples/source-gateway/server.mjs
```

Open `http://127.0.0.1:4320`, choose **Data > Connections & Refresh**, pick a source, then **Preview refresh** and **Apply reviewed refresh**. Edit `data/invoices.csv` and preview again to see the row count and control total change before anything is replaced.

Node prints an `ExperimentalWarning` for `node:sqlite`, which the SQLite connector uses.

## Check it

```sh
node examples/source-gateway/verify.mjs
```

```
PASS the dialog lists the gateway's two sources
PASS the preview shows the rows and the control total (4495.75)
PASS applying places the rows in a worksheet; apply checks access again
PASS a changed source previews as 4 → 2 rows with the control total change, then refreshes
PASS a value of the wrong type is refused as a schema change and the worksheet is kept
PASS no page errors
```

This is the check that the recovered editor and the newly written server agree: the dialog, its validation and its messages are the original client's; the gateway and connectors are new.
