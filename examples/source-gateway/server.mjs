/**
 * A host server with a source gateway: it serves the page and answers the editor's source
 * requests from a CSV file and a SQLite database.
 *
 *   node examples/source-gateway/build.mjs && node examples/source-gateway/server.mjs
 *   then open http://127.0.0.1:4320 and choose Data > Connections & Refresh.
 */
import { createServer } from "node:http";
import { readFile } from "node:fs/promises";
import { extname, join, normalize } from "node:path";
import { fileURLToPath } from "node:url";
import {
  createSourceGateway,
  createSourceGatewayHttpHandler,
} from "@ledgerglass/workbook/source-gateway";
import {
  createDelimitedFileConnector,
  createSqliteConnector,
} from "@ledgerglass/workbook/source-connectors";

const here = fileURLToPath(new URL(".", import.meta.url));

/** Creates the demonstration database. A real host already has its data. */
async function createLedgerDatabase(path) {
  const { DatabaseSync } = await import("node:sqlite");
  const database = new DatabaseSync(path);
  database.exec(`
    drop table if exists ledger;
    create table ledger (account text not null, posted text not null, amount real not null, memo text);
    insert into ledger values
      ('1000 Cash', '2026-01-31', 15250.75, 'Month-end balance'),
      ('1200 Receivables', '2026-01-31', 4495.75, null),
      ('2000 Payables', '2026-01-31', -3120.00, 'Supplier invoices'),
      ('4000 Revenue', '2026-01-31', -16626.50, null);
  `);
  database.close();
}

/**
 * @param options.dataDirectory  Folder holding invoices.csv; the demonstration database is created
 *   in it. Defaults to this example's data folder.
 */
export async function startServer({
  port = 4320,
  dataDirectory = join(here, "data"),
} = {}) {
  await createLedgerDatabase(join(dataDirectory, "ledger.db"));

  const gateway = createSourceGateway({
    connectors: [
      createDelimitedFileConnector({
        id: "invoices",
        label: "Open invoices (CSV)",
        directory: dataDirectory,
        file: "invoices.csv",
        columns: [
          { name: "Invoice", type: "text" },
          { name: "Date", type: "date" },
          { name: "Customer", type: "text" },
          { name: "Amount", type: "number" },
          { name: "Paid", type: "boolean" },
        ],
        // The editor totals control fields before and after a refresh so a reviewer can compare.
        controlFields: ["Amount"],
      }),
      createSqliteConnector({
        id: "ledger",
        label: "General ledger (SQLite)",
        database: join(dataDirectory, "ledger.db"),
        query:
          "select account, posted, amount, memo from ledger order by account",
        columns: [
          { name: "Account", type: "text" },
          { name: "Posted", type: "date" },
          { name: "Amount", type: "number" },
          { name: "Memo", type: "text", nullable: true },
        ],
        controlFields: ["Amount"],
      }),
    ],
    // Which caller may see which source. This example lets the one demo user see both.
    // A real host checks the user's engagement or role here.
    isAllowed: (sourceId, principal) => principal.user === "demo",
  });

  const handleSources = createSourceGatewayHttpHandler(gateway, {
    // Who is calling. A real host reads its session cookie or token here and returns null
    // when there is no valid session. This example treats every caller as one demo user.
    authenticate: () => ({ user: "demo" }),
  });

  const contentTypes = {
    ".html": "text/html; charset=utf-8",
    ".js": "text/javascript; charset=utf-8",
    ".css": "text/css; charset=utf-8",
  };
  const server = createServer(async (request, response) => {
    try {
      if (await handleSources(request, response)) return;
      const pathname = new URL(request.url, "http://localhost").pathname;
      const relative = normalize(pathname === "/" ? "/index.html" : pathname);
      if (request.method !== "GET" || relative.includes("..")) {
        response.writeHead(404).end();
        return;
      }
      const file = await readFile(join(here, "public", relative));
      response.writeHead(200, {
        "Content-Type":
          contentTypes[extname(relative)] ?? "application/octet-stream",
        // The editor starts its workers from blob: URLs.
        "Content-Security-Policy":
          "default-src 'self'; script-src 'self'; style-src 'self' 'unsafe-inline'; img-src 'self' data: blob:; font-src 'self' data:; worker-src 'self' blob:; connect-src 'self'",
      });
      response.end(file);
    } catch {
      response.writeHead(404).end();
    }
  });
  await new Promise((resolve) => server.listen(port, "127.0.0.1", resolve));
  return server;
}

if (process.argv[1] === fileURLToPath(import.meta.url)) {
  const server = await startServer({ port: Number(process.env.PORT ?? 4320) });
  console.log(`Open http://127.0.0.1:${server.address().port}`);
}
