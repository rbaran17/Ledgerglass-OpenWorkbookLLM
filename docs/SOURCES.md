# External sources

The editor's Data tab has a **Connections & Refresh** button. It opens a dialog where a person picks a source that your application offers, types a name for the output worksheet, previews the rows, and applies them. Later, **Refresh All** or a single connection's refresh reads the source again and repeats the preview-and-apply step.

The browser only ever names a source by its id. Your server decides who may see each source and how its data is read. The browser never sends a query, a file path, a URL or a credential.

## 1. What this is

The editor's side of this feature (the protocol, its limits, the checks on every reply, and the dialogs) was recovered from the compiled application. The server side was never in the browser bundle. The gateway and connectors in this library are therefore new code, written to that protocol. The repository's tests (`tests/unit/sources-gateway.test.js` and `tests/unit/sources-connectors.test.js`) run the editor's own catalog, authorize and refresh functions against the new gateway over a real HTTP server, so the replies are checked by the same rules the editor uses.

You need three pieces:

- **Server:** a gateway with one or more connectors, served on one path. See section 3.
- **Connectors:** the code that reads each source. This library includes five; you can write your own. See section 4.
- **Browser:** an adapter that sends the editor's requests to your server, passed to `AuditWorkbook`. See section 5.

## 2. The protocol: `ledgerglass-sources-v1`

Every request is an HTTP `POST` to your endpoint (default path `/api/sources`) with `Content-Type: application/json` and a body of at most 4,096 bytes by default (`maxBodyBytes` changes that). The body has exactly one of three shapes. Any other key or any other value is refused.

### Operations

**catalog** lists the sources the caller may see. Each one comes with an approval that says until when the editor may use it.

```json
{ "operation": "catalog" }
```

The reply is an array with one approval per visible source. This example shows one entry; a caller who may see both sources in section 3 gets both.

```json
[
  {
    "source": {
      "id": "invoices",
      "label": "Invoices",
      "kind": "file-folder",
      "columns": [
        { "name": "Invoice", "type": "text", "nullable": false },
        { "name": "Issued", "type": "date", "nullable": false },
        { "name": "Amount", "type": "number", "nullable": false }
      ],
      "controlFields": ["Amount"],
      "schemaVersion": "s-90c9a0db292d0240"
    },
    "expiresAt": "2026-11-12T21:26:35.812Z"
  }
]
```

**authorize** gives a fresh approval for one source. The editor calls it before each refresh.

```json
{ "operation": "authorize", "sourceId": "invoices" }
```

```json
{
  "source": {
    "id": "invoices",
    "label": "Invoices",
    "kind": "file-folder",
    "columns": [
      { "name": "Invoice", "type": "text", "nullable": false },
      { "name": "Issued", "type": "date", "nullable": false },
      { "name": "Amount", "type": "number", "nullable": false }
    ],
    "controlFields": ["Amount"],
    "schemaVersion": "s-90c9a0db292d0240"
  },
  "expiresAt": "2026-11-12T21:26:35.845Z"
}
```

**refresh** reads one source and returns its rows.

```json
{ "operation": "refresh", "sourceId": "invoices" }
```

```json
{
  "protocol": "ledgerglass-sources-v1",
  "source": {
    "id": "invoices",
    "label": "Invoices",
    "kind": "file-folder",
    "schemaVersion": "s-90c9a0db292d0240",
    "columns": [
      { "name": "Invoice", "type": "text", "nullable": false },
      { "name": "Issued", "type": "date", "nullable": false },
      { "name": "Amount", "type": "number", "nullable": false }
    ],
    "controlFields": ["Amount"]
  },
  "sourceTimestamp": "2026-10-10T08:00:00.000Z",
  "version": "v-b430705f5c9daa9168dbf3068fb5fa34",
  "rows": [
    ["INV-1001", "2026-10-02", 1200.5],
    ["INV-1002", "2026-10-05", 980],
    ["INV-1003", "2026-10-09", 45]
  ],
  "complete": true
}
```

The editor checks each reply itself. A refresh reply must describe exactly the source the approval named (same columns, control fields and schema version), and it must follow every rule below. A reply that breaks a rule is refused, and the workbook keeps its last good data.

### The source table

| Field           | Meaning                                                                                                                                     |
| --------------- | ------------------------------------------------------------------------------------------------------------------------------------------- |
| `id`            | The identifier you registered. Letters, digits, `.`, `-` and `_`; 1 to 100 characters, starting with a letter or digit.                     |
| `label`         | The name shown in the editor. Up to 128 characters, with no leading or trailing space and no control characters.                            |
| `kind`          | One of `database`, `odbc`, `odata`, `api`, `erp`, `file-folder`.                                                                            |
| `schemaVersion` | Changes exactly when the columns or control fields change. The gateway computes it unless the connector sets one.                           |
| `columns`       | Between 1 and 50 entries of `{ name, type, nullable }`. Names are unique, ignoring case.                                                    |
| `controlFields` | Names of numeric columns. The editor totals them before a refresh is applied, so a change in the source shows up as a change in the totals. |

### Column types and how each travels

Rows are arrays of values in column order, and each value must already have its column's type. The editor does no conversion of its own. The connectors and `shapeSourceRows` do the conversion described in section 4, before a reply is sent.

| Type      | JSON value                                                                                                            |
| --------- | --------------------------------------------------------------------------------------------------------------------- |
| `text`    | A string of at most 32,767 characters.                                                                                |
| `number`  | A finite JSON number. A connector given a `bigint` outside the safe integer range refuses it rather than rounding it. |
| `boolean` | `true` or `false`.                                                                                                    |
| `date`    | A string `"YYYY-MM-DD"` for a real calendar date, such as `"2026-10-05"`. Not a date-time.                            |

An empty value is `null`, and only a column declared `nullable: true` may hold it.

### Completeness

Every refresh reply carries `"complete": true`. The editor requires that exact value. A source that is too large is refused with the `limit` code. It is never cut short, so a reply never looks complete when it is not.

### Limits

These are the published limits. The gateway and the editor both enforce them.

| Limit                                          | Value                               | Applies to                                     |
| ---------------------------------------------- | ----------------------------------- | ---------------------------------------------- |
| Sources per gateway, and approvals per catalog | 20                                  | `createSourceGateway` and `catalog`            |
| Columns per source                             | 50                                  | Source description and each row                |
| Data rows per refresh                          | 5,000                               | Not counting the header                        |
| Cells per refresh                              | 100,000                             | `(rows + 1) × columns`, counting the header    |
| Text per cell                                  | 32,767 characters                   | Each `text` value                              |
| Response size                                  | 8,388,608 bytes (8 MiB)             | The reply; a file read by a file connector     |
| Time for one operation                         | 30 seconds                          | The gateway's default per-request limit        |
| Saved connection state in one workbook         | 200,000 characters, counted as JSON | The editor, across all sheets                  |
| Approval lifetime (gateway default)            | 15 minutes                          | `approvalTtlMs`, between 1 second and 24 hours |

### Errors

Every error reply has this body, with a status from the table below:

```json
{
  "error": {
    "code": "refused",
    "message": "This source is not authorized by the host. No source data was replaced."
  }
}
```

The editor shows its own text for each code, not the message in the body. For the codes in the table, the gateway's body message is the same text. Two replies carry their own messages: `401` says "Sign in to use this source.", and `415` says "Send application/json."

| Code          | Text the user sees                                                                                                                      | Gateway status                                                                                                                   |
| ------------- | --------------------------------------------------------------------------------------------------------------------------------------- | -------------------------------------------------------------------------------------------------------------------------------- |
| `unavailable` | The source is unavailable. The last successful data is retained. Check the host connection, then retry.                                 | 503                                                                                                                              |
| `refused`     | This source is not authorized by the host. No source data was replaced.                                                                 | 403; 401 when no caller can be identified                                                                                        |
| `expired`     | Source access expired. Renew access in the host application and preview again.                                                          | 403                                                                                                                              |
| `schema`      | The source schema changed. The working population is retained. Have the host approve the new schema and import it into a new worksheet. | 422                                                                                                                              |
| `limit`       | The source exceeds the published row, column, cell or response limit. No rows were truncated or published.                              | 413                                                                                                                              |
| `timeout`     | Source refresh timed out. The last successful data is retained; retry when the host is available.                                       | 504                                                                                                                              |
| `cancelled`   | Source refresh cancelled. The last successful data is retained.                                                                         | 499                                                                                                                              |
| `invalid`     | The source returned an invalid or incomplete result. The last successful data is retained.                                              | 400; 415 for a wrong content type. For a malformed request the gateway's message is "The request is not a valid source request." |
| `stale`       | The workbook, preview or source approval changed. Preview again before applying a refresh.                                              | 409                                                                                                                              |

The gateway raises `refused`, `unavailable`, `schema`, `limit`, `timeout`, `cancelled` and `invalid` itself, and passes through any code a connector raises. The editor raises `expired` when an approval's time has passed, and `stale` when a preview no longer matches the workbook. Other HTTP responses from the handler are: `401` when `authenticate` refuses the caller, `405` for any method other than `POST`, and `415` for a body that is not `application/json`.

## 3. Server

This complete example serves two sources from one Node server. It uses only `node:http`, the gateway and the connectors. `checkSession` is your own code: it returns the signed-in user, or `null`.

```js
import { createServer } from "node:http";
import {
  createSourceGateway,
  createSourceGatewayHttpHandler,
} from "@ledgerglass/workbook/source-gateway";
import {
  createDelimitedFileConnector,
  createSqliteConnector,
} from "@ledgerglass/workbook/source-connectors";

const invoices = createDelimitedFileConnector({
  id: "invoices",
  label: "Invoices",
  directory: "/srv/exports",
  file: "invoices.csv",
  columns: [
    { name: "Invoice", type: "text" },
    { name: "Issued", type: "date" },
    { name: "Amount", type: "number" },
  ],
  controlFields: ["Amount"],
});

const openItems = createSqliteConnector({
  id: "open-items",
  label: "Open items",
  database: "/srv/ledger/ledger.db",
  query: "SELECT account, due, balance FROM open_items WHERE balance <> 0",
  columns: [
    { name: "Account", type: "text" },
    { name: "Due", type: "date", nullable: true },
    { name: "Balance", type: "number" },
  ],
  controlFields: ["Balance"],
});

const gateway = createSourceGateway({
  connectors: [invoices, openItems],
  // Each caller sees only the sources its roles allow.
  isAllowed: (sourceId, user) => user.roles.includes("finance"),
});

const handleSources = createSourceGatewayHttpHandler(gateway, {
  // checkSession is your own code: it returns the signed-in user, or null.
  authenticate: async (request) => checkSession(request.headers.cookie),
});

createServer(async (request, response) => {
  if (await handleSources(request, response)) return;
  response.writeHead(404).end();
}).listen(8080);
```

What each part does:

- **`createSourceGateway({ connectors, isAllowed })`** fixes the list of sources when it is created. It checks each connector's description and throws `TypeError` if one is wrong. The browser cannot add or change a source.
- **`authenticate`** answers who is calling. It receives the request and returns the caller's identity, which the gateway passes unchanged to `isAllowed`. Return `null` to refuse the request with `401`. If it throws, that counts as `null`.
- **`isAllowed(sourceId, principal)`** answers whether this caller may see this source. It can return a boolean or a promise of one. If it throws, the answer is no.
- **Neither has a default that allows access.** Both are required, and `createSourceGateway` and `createSourceGatewayHttpHandler` throw `TypeError` if either is missing. Pass `() => true` only for a server with no users to tell apart.
- **An unknown source and a forbidden source get the same reply:** `403` with the `refused` code and the same message. The caller cannot learn which sources exist.

The handler answers only `POST` on its path and returns `false` for any other path, so you can serve other routes from the same server. A runnable version that drives the editor's dialog is in `examples/source-gateway` in the repository.

## 4. Connectors

The examples in this section assume this import:

```js
import {
  createDelimitedFileConnector,
  createJsonApiConnector,
  createODataConnector,
  createQueryConnector,
  createSqliteConnector,
} from "@ledgerglass/workbook/source-connectors";
```

Each connector is an object with an `id`, `label`, `kind`, `columns`, optional `controlFields`, and a `read` function. Every connector takes `id`, `label`, `columns` and `controlFields`, and the options listed below. Everything a connector reaches (a folder, a URL, a database file, a query) is fixed when it is created.

Values are typed strictly. A value that does not match its column's type is a `schema` error. The one exception is text sources such as CSV, where the text is read as the column's type using the rules in the file connector's notes.

### `createDelimitedFileConnector`

Reads one CSV or TSV file inside a folder. Its `kind` is `file-folder`.

| Option      | Required | Meaning                                                                                                            |
| ----------- | -------- | ------------------------------------------------------------------------------------------------------------------ |
| `directory` | Yes      | The folder. The file must resolve inside it, also through links.                                                   |
| `file`      | Yes      | The file name inside the folder. Not absolute, and it may not contain `..`.                                        |
| `delimiter` | No       | `","`, `"\t"` or `";"` to fix the delimiter. Default: detected.                                                    |
| `hasHeader` | No       | Default `true`. The header row must match the column names, in order. With `false`, columns are taken by position. |

Rules for the file: a leading byte-order mark is ignored, blank lines are skipped, and every data row must have exactly as many fields as there are columns. Text is read as the column's type: `"1200.5"` is a number, `"true"` or `"false"` is a boolean, and an empty field is `null`. A date must be `YYYY-MM-DD` or a date-time whose date part is real. The `sourceTimestamp` is the file's modified time.

```js
const invoices = createDelimitedFileConnector({
  id: "invoices",
  label: "Invoices",
  directory: "/srv/exports",
  file: "invoices.csv",
  columns: [
    { name: "Invoice", type: "text" },
    { name: "Issued", type: "date" },
    { name: "Amount", type: "number" },
  ],
  controlFields: ["Amount"],
});
```

### `createJsonApiConnector`

Reads a JSON endpoint that returns a list of records. Its `kind` is `api` unless you set one.

| Option     | Required | Meaning                                                                                             |
| ---------- | -------- | --------------------------------------------------------------------------------------------------- |
| `url`      | Yes      | The endpoint, as an absolute `http` or `https` URL. Fixed here.                                     |
| `headers`  | No       | Sent with each request, for example an API key. Stays on the server.                                |
| `rowsPath` | No       | Dot path to the list in the reply, such as `"data.items"`. Default: the reply is the list.          |
| `fields`   | No       | The record key for each column, in column order. Dot paths are followed. Default: the column names. |
| `kind`     | No       | One of the source kinds. Default `"api"`.                                                           |
| `fetch`    | No       | The fetch function to use. Default: the global `fetch`.                                             |

Requests are `GET` with `Accept: application/json`, `redirect: "error"`, and the operation's signal. A redirect therefore fails the read as `unavailable`; the connector does not follow it. A `401` or `403` reply is `refused`, another failed status is `unavailable`, a body that is not JSON is `invalid`, and a body over 8 MiB is `limit`.

```js
const orders = createJsonApiConnector({
  id: "orders",
  label: "Orders",
  url: "https://api.example.com/v2/orders",
  headers: { "X-Api-Key": process.env.ORDERS_API_KEY },
  rowsPath: "data.items",
  fields: ["orderNo", "customer.name", "total"],
  columns: [
    { name: "Order", type: "text" },
    { name: "Customer", type: "text" },
    { name: "Total", type: "number" },
  ],
  controlFields: ["Total"],
});
```

### `createODataConnector`

Reads an OData entity set. Its `kind` is `odata`.

| Option    | Required | Meaning                                                                                 |
| --------- | -------- | --------------------------------------------------------------------------------------- |
| `url`     | Yes      | The entity set URL, for example `"https://erp.example.com/odata/Invoices"`. Fixed here. |
| `filter`  | No       | A fixed `$filter` expression, sent with every page request.                             |
| `fields`  | No       | The property name for each column, in column order. Default: the column names.          |
| `headers` | No       | Sent with each request. Stays on the server.                                            |
| `fetch`   | No       | The fetch function to use. Default: the global `fetch`.                                 |

The connector asks only for the declared properties (`$select`) and follows `@odata.nextLink` pages. A page link on any other origin than `url` is refused and never fetched. A set of more than 200 pages or 5,000 rows is refused with `limit`.

```js
const erpInvoices = createODataConnector({
  id: "erp-invoices",
  label: "ERP invoices",
  url: "https://erp.example.com/odata/Invoices",
  filter: "Status eq 'open'",
  fields: ["InvoiceNumber", "Total"],
  headers: { Authorization: `Basic ${process.env.ERP_CREDENTIALS}` },
  columns: [
    { name: "Number", type: "text" },
    { name: "Total", type: "number" },
  ],
});
```

### `createSqliteConnector`

Runs one fixed `SELECT` on a SQLite file. Its `kind` is `database`. It uses Node's built-in `node:sqlite`, which needs Node 22.5 or later.

| Option       | Required | Meaning                                                                                             |
| ------------ | -------- | --------------------------------------------------------------------------------------------------- |
| `database`   | Yes      | Path to the database file. It is opened read-only.                                                  |
| `query`      | Yes      | One `SELECT` or `WITH ... SELECT` statement. Its result columns are the declared columns, in order. |
| `parameters` | No       | Values for the `?` placeholders in the query. Fixed here.                                           |

The query is checked when the connector is created. A statement with a second statement after a semicolon, or with a write keyword (such as `INSERT`, `UPDATE`, `DELETE`, `DROP`, `CREATE`, `ATTACH` or `PRAGMA`) outside a string or a quoted name, throws `TypeError`. The check is only a first screen. The read-only connection is what stops writes.

```js
const openItems = createSqliteConnector({
  id: "open-items",
  label: "Open items",
  database: "/srv/ledger/ledger.db",
  query: "SELECT account, balance FROM open_items WHERE balance <> ?",
  parameters: [0],
  columns: [
    { name: "Account", type: "text" },
    { name: "Balance", type: "number" },
  ],
  controlFields: ["Balance"],
});
```

### `createQueryConnector`

Runs a fixed query through a function you supply. Use it for any database driver, an ODBC connection or an ERP client. The driver stays in your application, and this library does not depend on one. Its `kind` is `database` unless you set `"odbc"` or `"erp"`.

| Option | Required | Meaning                                                                                                                      |
| ------ | -------- | ---------------------------------------------------------------------------------------------------------------------------- |
| `run`  | Yes      | `({ signal }) => records`, or a promise of them. Each record is an array in column order, or an object keyed by column name. |
| `kind` | No       | `"database"` (default), `"odbc"` or `"erp"`.                                                                                 |

An error from `run` keeps its code if it is an `ExternalSourceError`. An abort or timeout maps to `cancelled` or `timeout`. Any other error is `unavailable`. Build the query inside `run`; nothing from a request should reach it.

```js
const stock = createQueryConnector({
  id: "stock-on-hand",
  label: "Stock on hand",
  kind: "erp",
  columns: [
    { name: "Item", type: "text" },
    { name: "OnHand", type: "number" },
  ],
  // erp is your own client, opened in your application. The query is fixed here.
  async run({ signal }) {
    const rows = await erp.query(
      "SELECT item, on_hand FROM stock WHERE on_hand > 0",
      { signal },
    );
    return rows.map((row) => [row.item, row.on_hand]);
  },
});
```

### Writing your own connector

A connector is a plain object with this shape:

```js
const connector = {
  id, // letters, digits, dots, dashes or underscores
  label, // the name shown in the editor
  kind, // one of the six source kinds
  columns, // [{ name, type, nullable? }], 1 to 50 columns
  controlFields, // optional: numeric column names
  async read({ signal }) {
    // returns { rows, sourceTimestamp?, version? }
  },
};
```

`read` resolves to `{ rows, sourceTimestamp?, version? }`. `sourceTimestamp` is optional and accepts anything `new Date(...)` takes; by default it is the time of the read. `version` is optional and must be a valid identifier; by default the gateway hashes the rows. Build the rows with `shapeSourceRows`, which checks each value against its column's type and the row and cell limits:

```js
import {
  ExternalSourceError,
  shapeSourceRows,
} from "@ledgerglass/workbook/source-gateway";

const columns = [
  { name: "Code", type: "text" },
  { name: "Rate", type: "number" },
];

const rates = {
  id: "fx-rates",
  label: "FX rates",
  kind: "api",
  columns,
  async read({ signal }) {
    const response = await fetch("https://rates.example.com/today", {
      signal,
      redirect: "error",
    });
    if (!response.ok) throw new ExternalSourceError("unavailable");
    const body = await response.json();
    return {
      rows: shapeSourceRows(body.rates, columns),
      sourceTimestamp: body.asOf,
    };
  },
};
```

`shapeSourceRows(records, columns, options)` takes an array of records, each an array in column order or an object keyed by column name. It throws `ExternalSourceError` with `limit` over the row or cell limit, and with `schema` for a value of the wrong type. Pass `{ parseText: true }` only when the values are text that should be read as the column's type.

Throw `ExternalSourceError` with a code for a failure you expect, so the user sees the right message. Any other error reaches the user as `unavailable`.

## 5. Browser

`createHttpSourceAdapter` sends the editor's requests to your endpoint. Pass its result as the `externalSourceAdapter` prop of `AuditWorkbook`:

```jsx
import { AuditWorkbook, createHttpSourceAdapter } from "@ledgerglass/workbook";
import "@ledgerglass/workbook/style.css";

const externalSourceAdapter = createHttpSourceAdapter({
  endpoint: "/api/sources",
});

export function Ledger() {
  return (
    <div style={{ height: 640 }}>
      <AuditWorkbook externalSourceAdapter={externalSourceAdapter} />
    </div>
  );
}
```

| Option        | Required | Meaning                                                        |
| ------------- | -------- | -------------------------------------------------------------- |
| `endpoint`    | Yes      | The URL of your gateway, such as `"/api/sources"`.             |
| `headers`     | No       | Sent with each request, for example an `Authorization` header. |
| `fetch`       | No       | The fetch function to use. Default: the global `fetch`.        |
| `credentials` | No       | Credential handling for the request. Default `"same-origin"`.  |

The adapter has three methods, `catalog`, `authorize` and `refresh`. Each returns the parsed reply, and the editor checks it again before using it. A network failure becomes `unavailable`, an abort becomes `cancelled`, and a reply with an error code keeps that code. A failed status with no known code becomes `refused` for `401` and `403`, and `unavailable` for anything else. `createHttpSourceAdapter` throws `TypeError` if `endpoint` is missing.

## 6. Security

What the code does:

- Sources are fixed when the gateway is created. A request carries only an operation and, for two of the three operations, one source id. Any other key is refused.
- The file connector resolves links before it checks the folder, and it refuses any file that resolves outside the folder. A file name containing `..` or an absolute path is refused when the connector is created.
- HTTP connectors refuse redirects. OData page links stay on the origin you configured; any other origin is refused and never fetched.
- SQLite is opened read-only. The query is one `SELECT` (or `WITH ... SELECT`), checked when the connector is created.
- Values are typed strictly and are never guessed. Text sources are read by the rules in section 4. Outside those rules, the only conversions are a `Date` becoming its `YYYY-MM-DD` date and a `bigint` within the safe integer range becoming a number.
- Data over a limit is refused, never truncated. Nothing reaches the workbook from a partial read.
- A source the caller may not see and a source that does not exist get the same reply, so the caller cannot tell which sources exist.
- Every JSON reply is sent with `Cache-Control: no-store` and `X-Content-Type-Options: nosniff`.

What your host must still do:

- **Authentication.** `authenticate` decides who is calling. Protect the endpoint the same way as the rest of your application, and check the session on every request.
- **Transport security.** Serve the endpoint over HTTPS. Credentials in cookies or headers travel with every request.
- **Rate limiting.** The gateway does not limit how often a caller may refresh. A caller can refresh a source as often as it likes, up to the limits above.
- **Auditing.** The gateway does not keep a log. Record who refreshed which source and when, in your own logs.
- **Credentials and connection details.** Keep API keys, passwords, connection strings and database paths on the server. The browser never receives them, but your connectors hold them, so keep credential files and the database file readable only by the server process. Give each database login only the rights the query needs.
- **Folder permissions.** Keep the exports folder readable by the server process only, and outside any path your web server publishes.

## 7. Limits of this implementation

- **No connection pooling and no caching.** Each refresh opens the file, the database or the endpoint again.
- **No incremental refresh.** A refresh sends the whole table, so a source is limited to 5,000 rows and 100,000 cells.
- **SQLite support is experimental.** `node:sqlite` is marked experimental by Node, and it prints a warning when loaded.
- **No ODBC driver is included.** Use `createQueryConnector` with a driver your application installs.
- **Not tested against a live OData service or ERP.** The OData connector was tested against a stand-in server. It sends `$select` and `$filter` with a literal `$` and spaces as `%20`.
- **Approvals are not stored.** The gateway does not remember which approvals it issued, and a refresh checks only that the caller may see the source. The `expiresAt` time is advisory to the client: the editor refuses an expired approval, but a client that ignores the time can still call `refresh`.
- **Timestamps default to the read time.** Connectors that do not report a modification time (the JSON, OData and query connectors) report the time of the read.
