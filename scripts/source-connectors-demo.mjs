/** Local synthetic acceptance server, NOT a production authentication or database deployment. */
import { createServer } from "node:http";
import { randomUUID } from "node:crypto";
import { mkdtemp, mkdir, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import { DatabaseSync } from "node:sqlite";
import {
  createJsonApiSource,
  createODataSource,
  createErpSource,
  createFolderSource,
  createSqliteSource,
} from "../lib/source-connectors.js";
import {
  createExternalSourceGateway,
  createExternalSourceHttpHandler,
} from "../lib/source-gateway.js";

export async function startConnectedSourceDemo({ port = 4182 } = {}) {
  const directory = await mkdtemp(
      join(tmpdir(), "ledgerglass-connected-demo-"),
    ),
    folder = join(directory, "csv");
  await mkdir(folder);
  const dbPath = join(directory, "receivables.sqlite"),
    session = randomUUID(),
    upstreamKey = randomUUID();
  const columns = [
    { name: "ID", type: "text", nullable: false },
    { name: "Amount", type: "number", nullable: false },
    { name: "Date", type: "date", nullable: false },
  ];
  let version = 1,
    expired = false,
    handler,
    origin;
  const records = () => [
    { ID: "00123", Amount: 100, Date: "2026-09-01" },
    { ID: "00124", Amount: 200, Date: "2026-09-02" },
    ...(version === 2 ? [{ ID: "00125", Amount: 50, Date: "2026-09-03" }] : []),
  ];
  async function publish() {
    const db = new DatabaseSync(dbPath);
    try {
      db.exec(
        "BEGIN; CREATE TABLE IF NOT EXISTS receivables (ID TEXT, Amount REAL, Date TEXT); DELETE FROM receivables",
      );
      const insert = db.prepare("INSERT INTO receivables VALUES (?,?,?)");
      for (const r of records()) insert.run(r.ID, r.Amount, r.Date);
      db.exec("COMMIT");
    } finally {
      db.close();
    }
    await writeFile(
      join(folder, "receivables.csv"),
      "ID,Amount,Date\n" +
        records()
          .map((r) => `${r.ID},${r.Amount},${r.Date}`)
          .join("\n") +
        "\n",
    );
  }
  await publish();
  const server = createServer(async (req, res) => {
    const reply = (data, status = 200, headers = {}) => {
      res.writeHead(status, {
        "Content-Type": "application/json",
        "Cache-Control": "no-store",
        "X-Content-Type-Options": "nosniff",
        ...headers,
      });
      res.end(JSON.stringify(data));
    };
    try {
      if (!origin || req.headers.host !== new URL(origin).host)
        return reply({ error: "refused" }, 403);
      const url = new URL(req.url ?? "/", origin);
      if (
        (url.pathname === "/" ||
          url.pathname === "/external-connectors.html") &&
        req.method === "GET"
      ) {
        res.writeHead(200, {
          "Content-Type": "text/html; charset=utf-8",
          "Cache-Control": "no-store",
          "Set-Cookie": `lg-source-demo=${session}; HttpOnly; SameSite=Strict; Path=/`,
        });
        return res.end(
          await readFile(
            new URL("../dist/external-connectors.html", import.meta.url),
          ),
        );
      }
      if (
        ["/fixture/api", "/fixture/odata", "/fixture/erp"].includes(
          url.pathname,
        )
      ) {
        if (
          req.method !== "GET" ||
          req.headers.authorization !== `Bearer ${upstreamKey}` ||
          expired
        )
          return reply({ error: "expired" }, 401);
        const current = Number(
            url.searchParams.get("page") ?? url.searchParams.get("$skip") ?? 1,
          ),
          all = records(),
          row = all[current - 1];
        const isOData = url.pathname !== "/fixture/api",
          more = current < all.length;
        const next = `?tenant=demo&${isOData ? "$skip" : "page"}=${current + 1}`;
        return reply(
          isOData
            ? {
                value: row ? [row] : [],
                "@odata.count": all.length,
                ...(more ? { "@odata.nextLink": next } : {}),
              }
            : {
                rows: row ? [row] : [],
                count: all.length,
                sourceTimestamp: "2026-09-18T09:00:00Z",
                version: `v${version}`,
                ...(more ? { nextLink: next } : {}),
              },
          200,
          {
            "x-source-time": "2026-09-18T09:00:00Z",
            "x-source-version": `v${version}`,
          },
        );
      }
      if (req.headers.cookie !== `lg-source-demo=${session}`)
        return reply({ error: "expired" }, 401);
      if (req.method !== "POST" || req.headers.origin !== origin)
        return reply({ error: "refused" }, 403);
      if (
        ["/demo/advance", "/demo/expire", "/demo/restore"].includes(
          url.pathname,
        )
      ) {
        if (url.pathname === "/demo/advance") {
          version = 2;
          await publish();
        } else expired = url.pathname === "/demo/expire";
        return reply({
          message: expired
            ? "API access expired; previous workbook data is retained on refresh failure"
            : `Version ${version}: ${records().length} records, total ${version === 1 ? 300 : 350}`,
        });
      }
      if (url.pathname !== "/api/workbook-sources")
        return reply({ error: "refused" }, 404);
      const chunks = [];
      let bytes = 0;
      for await (const chunk of req) {
        bytes += chunk.length;
        if (bytes > 1024) return reply({ error: "limit" }, 413);
        chunks.push(chunk);
      }
      const result = await handler(
        new Request(url, {
          method: "POST",
          headers: {
            "content-type": req.headers["content-type"] ?? "",
            origin,
            cookie: req.headers.cookie,
          },
          body: Buffer.concat(chunks),
        }),
      );
      res.writeHead(result.status, Object.fromEntries(result.headers));
      res.end(Buffer.from(await result.arrayBuffer()));
    } catch {
      if (!res.headersSent) reply({ error: "unavailable" }, 503);
      else res.end();
    }
  });
  await new Promise((resolve, reject) => {
    server.once("error", reject);
    server.listen(port, "127.0.0.1", resolve);
  });
  origin = `http://127.0.0.1:${server.address().port}`;
  const descriptor = (id, label, kind) => ({
    id,
    label,
    kind,
    schemaVersion: "v1",
    columns,
    controlFields: ["Amount"],
  });
  const headers = async () => ({ Authorization: `Bearer ${upstreamKey}` }),
    snapshotHeaders = {
      timestamp: "x-source-time",
      version: "x-source-version",
    };
  const sources = [
    createJsonApiSource({
      source: descriptor("demo.api", "Local paginated API", "api"),
      endpoint: `${origin}/fixture/api?tenant=demo`,
      headers,
    }),
    createODataSource({
      source: descriptor("demo.odata", "Local OData service", "odata"),
      endpoint: `${origin}/fixture/odata?tenant=demo`,
      headers,
      snapshotHeaders,
    }),
    createErpSource({
      source: descriptor("demo.erp", "Local ERP contract", "erp"),
      endpoint: `${origin}/fixture/erp?tenant=demo`,
      headers,
      snapshotHeaders,
      transport: "odata",
    }),
    createSqliteSource({
      source: descriptor(
        "demo.sqlite",
        "Read-only SQLite database",
        "database",
      ),
      databasePath: dbPath,
      query: "SELECT ID,Amount,Date FROM receivables ORDER BY ID",
    }),
    createFolderSource({
      source: descriptor("demo.folder", "Local CSV folder", "file-folder"),
      directory: folder,
      format: "csv",
    }),
  ];
  const gateway = createExternalSourceGateway({
    sources,
    authorize: async (principal) =>
      principal === "local-demo-user"
        ? { expiresAt: new Date(Date.now() + 60000).toISOString() }
        : null,
  });
  handler = createExternalSourceHttpHandler({
    gateway,
    authenticate: async (request) =>
      request.headers.get("cookie") === `lg-source-demo=${session}`
        ? "local-demo-user"
        : null,
    allowRequest: (request) => request.headers.get("origin") === origin,
  });
  return {
    origin,
    directory,
    async close() {
      server.closeAllConnections();
      await new Promise((resolve) => server.close(resolve));
      await rm(directory, { recursive: true, force: true });
    },
  };
}
if (process.argv[1] === fileURLToPath(import.meta.url)) {
  const demo = await startConnectedSourceDemo({
    port: Number(process.env.PORT ?? 4182),
  });
  console.log(
    `Synthetic connector demo: ${demo.origin}/external-connectors.html`,
  );
  for (const signal of ["SIGINT", "SIGTERM"])
    process.once(signal, async () => {
      await demo.close();
      process.exit(0);
    });
}
