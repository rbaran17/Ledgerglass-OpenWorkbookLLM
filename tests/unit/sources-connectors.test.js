import { afterAll, beforeAll, describe, expect, test } from "vitest";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import {
  createDelimitedFileConnector,
  createJsonApiConnector,
  createODataConnector,
  createQueryConnector,
  createSqliteConnector,
} from "../../src/sources/connectors.js";
import {
  createSourceGateway,
  ExternalSourceError,
  externalSourceLimits,
} from "../../src/sources/gateway.js";

const { DatabaseSync } = await import("node:sqlite");

/** Resolves to the error a rejected promise carries; fails the test if it resolves. */
async function failureOf(promise) {
  try {
    await promise;
  } catch (error) {
    return error;
  }
  throw new Error("expected the promise to reject");
}

function expectSourceError(error, code) {
  expect(error).toBeInstanceOf(ExternalSourceError);
  expect(error.code).toBe(code);
}

async function makeFolder() {
  return fs.mkdtemp(path.join(os.tmpdir(), "lg-src-"));
}

/** A fake fetch: returns the given replies in order; an Error reply is thrown instead. */
function fakeFetch(...replies) {
  const calls = [];
  const fetchImplementation = async (url, init) => {
    calls.push({ url: String(url), init });
    const next = replies.shift();
    if (next instanceof Error) throw next;
    return next;
  };
  fetchImplementation.calls = calls;
  return fetchImplementation;
}

function reply(body, init = {}) {
  return new Response(
    typeof body === "string" ? body : JSON.stringify(body),
    init,
  );
}

const textNumber = [
  { name: "item", type: "text" },
  { name: "qty", type: "number" },
];

const typedColumns = [
  { name: "item", type: "text" },
  { name: "qty", type: "number" },
  { name: "paid", type: "boolean" },
  { name: "due", type: "date" },
];

describe("createDelimitedFileConnector", () => {
  let folder;
  beforeAll(async () => {
    folder = await makeFolder();
  });
  afterAll(async () => {
    await fs.rm(folder, { recursive: true, force: true });
  });

  async function write(name, text) {
    const full = path.join(folder, name);
    await fs.mkdir(path.dirname(full), { recursive: true });
    await fs.writeFile(full, text);
    return full;
  }

  const connector = (file, columns = typedColumns, extra = {}) =>
    createDelimitedFileConnector({
      id: "orders-file",
      label: "Orders file",
      directory: folder,
      file,
      columns,
      ...extra,
    });

  test("reads a CSV with a header into typed rows", async () => {
    await write(
      "typed.csv",
      "item,qty,paid,due\nwidget,3,true,2026-01-02\ngadget,1.5,false,2026-02-03\n",
    );
    const result = await connector("typed.csv").read();
    expect(result.rows).toEqual([
      ["widget", 3, true, "2026-01-02"],
      ["gadget", 1.5, false, "2026-02-03"],
    ]);
  });

  test("has kind file-folder and keeps its id, label and columns", async () => {
    const built = connector("typed.csv");
    expect(built.kind).toBe("file-folder");
    expect(built.id).toBe("orders-file");
    expect(built.label).toBe("Orders file");
    expect(built.columns).toBe(typedColumns);
  });

  test("sourceTimestamp is the file's modified time", async () => {
    const full = await write("stamped.csv", "item,qty\nwidget,3\n");
    const when = new Date("2025-01-02T03:04:05.000Z");
    await fs.utimes(full, when, when);
    const result = await connector("stamped.csv", textNumber).read();
    expect(result.sourceTimestamp).toEqual(when);
  });

  test("an empty field is null and refused in a non-nullable column", async () => {
    await write(
      "empty-required.csv",
      "item,qty,paid,due\n,1,true,2026-01-02\n",
    );
    const error = await failureOf(connector("empty-required.csv").read());
    expectSourceError(error, "schema");
  });

  test("an empty field is accepted as null in a nullable column", async () => {
    await write("empty-nullable.csv", "item,qty\n,1\nwidget,\n");
    const columns = [
      { name: "item", type: "text", nullable: true },
      { name: "qty", type: "number", nullable: true },
    ];
    const result = await connector("empty-nullable.csv", columns).read();
    expect(result.rows).toEqual([
      [null, 1],
      ["widget", null],
    ]);
  });

  test("a header that does not match the declared columns is refused", async () => {
    await write("order.csv", "qty,item,paid,due\n1,widget,true,2026-01-02\n");
    await write(
      "wrong-name.csv",
      "item,quantity,paid,due\nwidget,1,true,2026-01-02\n",
    );
    await write(
      "extra.csv",
      "item,qty,paid,due,note\nwidget,1,true,2026-01-02,x\n",
    );
    for (const name of ["order.csv", "wrong-name.csv", "extra.csv"]) {
      const error = await failureOf(connector(name).read());
      expectSourceError(error, "schema");
    }
  });

  test("a data row with too few or too many fields is refused", async () => {
    await write("short-row.csv", "item,qty,paid,due\nwidget,1,true\n");
    await write(
      "long-row.csv",
      "item,qty,paid,due\nwidget,1,true,2026-01-02,extra\n",
    );
    expectSourceError(
      await failureOf(connector("short-row.csv").read()),
      "schema",
    );
    expectSourceError(
      await failureOf(connector("long-row.csv").read()),
      "schema",
    );
  });

  test("text in a number column is refused", async () => {
    await write("text-in-number.csv", "item,qty\nwidget,three\n");
    const error = await failureOf(
      connector("text-in-number.csv", textNumber).read(),
    );
    expectSourceError(error, "schema");
  });

  test("hasHeader false reads the columns by position", async () => {
    await write("no-header.csv", "widget,3,true,2026-01-02\n");
    const result = await connector("no-header.csv", typedColumns, {
      hasHeader: false,
    }).read();
    expect(result.rows).toEqual([["widget", 3, true, "2026-01-02"]]);
  });

  test("an explicit semicolon delimiter is used", async () => {
    await write("semicolon.csv", "item;qty\nwidget;2\n");
    const result = await connector("semicolon.csv", textNumber, {
      delimiter: ";",
    }).read();
    expect(result.rows).toEqual([["widget", 2]]);
  });

  test("a tab-separated file is read", async () => {
    await write("tabs.tsv", "item\tqty\nwidget\t2\n");
    const result = await connector("tabs.tsv", textNumber).read();
    expect(result.rows).toEqual([["widget", 2]]);
  });

  test("a UTF-8 byte order mark is ignored", async () => {
    await write("bom.csv", "﻿item,qty\nwidget,2\n");
    const result = await connector("bom.csv", textNumber).read();
    expect(result.rows).toEqual([["widget", 2]]);
  });

  test("quoted fields keep commas, and doubled quotes become one quote", async () => {
    await write("quoted.csv", 'item,qty\n"bolt, large ""M8""",4\n');
    const result = await connector("quoted.csv", textNumber).read();
    expect(result.rows).toEqual([['bolt, large "M8"', 4]]);
  });

  test("blank lines are skipped", async () => {
    await write("blank-lines.csv", "item,qty\n\nwidget,2\n\n\ngadget,3\n\n");
    const result = await connector("blank-lines.csv", textNumber).read();
    expect(result.rows).toEqual([
      ["widget", 2],
      ["gadget", 3],
    ]);
  });

  test("a missing file is an ExternalSourceError with code unavailable", async () => {
    const error = await failureOf(connector("not-there.csv").read());
    expectSourceError(error, "unavailable");
  });

  test("construction throws TypeError for a missing directory or file option", () => {
    expect(() =>
      createDelimitedFileConnector({ columns: textNumber, file: "a.csv" }),
    ).toThrow(TypeError);
    expect(() =>
      createDelimitedFileConnector({ columns: textNumber, directory: folder }),
    ).toThrow(TypeError);
  });

  test("construction throws TypeError for an absolute file or a file containing ..", () => {
    for (const file of ["/etc/passwd", "../outside.csv", "sub/../a.csv"]) {
      expect(() =>
        createDelimitedFileConnector({
          directory: folder,
          file,
          columns: textNumber,
        }),
      ).toThrow(TypeError);
    }
  });

  test("a symbolic link inside the folder that points outside it is refused", async () => {
    const outside = path.join(
      path.dirname(folder),
      `${path.basename(folder)}-outside.csv`,
    );
    await fs.writeFile(outside, "item,qty\nsecret,9\n");
    await fs.symlink(outside, path.join(folder, "escape.csv"));
    try {
      const error = await failureOf(connector("escape.csv", textNumber).read());
      expectSourceError(error, "refused");
    } finally {
      await fs.rm(outside, { force: true });
    }
  });

  test("a symbolic link inside the folder that points to a file inside it is read", async () => {
    await write("target.csv", "item,qty\nwidget,2\n");
    await fs.symlink(
      path.join(folder, "target.csv"),
      path.join(folder, "inner-link.csv"),
    );
    const result = await connector("inner-link.csv", textNumber).read();
    expect(result.rows).toEqual([["widget", 2]]);
  });

  test("a sub-folder given as file is not read", async () => {
    await fs.mkdir(path.join(folder, "sub-folder"), { recursive: true });
    const error = await failureOf(connector("sub-folder", textNumber).read());
    expectSourceError(error, "unavailable");
  });

  test("a file larger than the response byte limit is refused with code limit", async () => {
    const big = path.join(folder, "big.csv");
    const body =
      "item\n" +
      "1\n".repeat(Math.ceil(externalSourceLimits.responseBytes / 2) + 1);
    await fs.writeFile(big, body);
    const error = await failureOf(
      connector("big.csv", [{ name: "item", type: "number" }]).read(),
    );
    expectSourceError(error, "limit");
  });

  test("more rows than the row limit is refused with code limit", async () => {
    await write(
      "many.csv",
      "item\n" + "1\n".repeat(externalSourceLimits.rows + 1),
    );
    const error = await failureOf(
      connector("many.csv", [{ name: "item", type: "number" }]).read(),
    );
    expectSourceError(error, "limit");
  });
});

describe("createJsonApiConnector", () => {
  const columns = [
    { name: "item", type: "text" },
    { name: "qty", type: "number" },
  ];
  const url = "https://api.example.test/items";

  test("reads a top-level list of records", async () => {
    const fetch = fakeFetch(reply([{ item: "widget", qty: 3 }]));
    const result = await createJsonApiConnector({
      id: "items-api",
      label: "Items API",
      url,
      columns,
      fetch,
    }).read();
    expect(result.rows).toEqual([["widget", 3]]);
  });

  test("rowsPath and dotted field paths select the list and its values", async () => {
    const fetch = fakeFetch(
      reply({ data: { items: [{ name: "widget", stock: { qty: 3 } }] } }),
    );
    const result = await createJsonApiConnector({
      url,
      rowsPath: "data.items",
      fields: ["name", "stock.qty"],
      columns,
      fetch,
    }).read();
    expect(result.rows).toEqual([["widget", 3]]);
  });

  test("kind defaults to api, and a custom kind is kept", () => {
    const base = { url, columns, fetch: fakeFetch() };
    expect(createJsonApiConnector(base).kind).toBe("api");
    expect(createJsonApiConnector({ ...base, kind: "erp" }).kind).toBe("erp");
  });

  test("sends Accept and the custom headers with GET, redirect error and the signal", async () => {
    const fetch = fakeFetch(reply([]));
    const controller = new AbortController();
    await createJsonApiConnector({
      url,
      columns,
      headers: { Authorization: "Bearer test-token" },
      fetch,
    }).read({ signal: controller.signal });
    const [call] = fetch.calls;
    expect(call.url).toBe(url);
    expect(call.init.method).toBe("GET");
    expect(call.init.redirect).toBe("error");
    expect(call.init.headers).toEqual({
      Accept: "application/json",
      Authorization: "Bearer test-token",
    });
    expect(call.init.signal).toBe(controller.signal);
  });

  test("401 and 403 replies are refused", async () => {
    for (const status of [401, 403]) {
      const error = await failureOf(
        createJsonApiConnector({
          url,
          columns,
          fetch: fakeFetch(reply("{}", { status })),
        }).read(),
      );
      expectSourceError(error, "refused");
    }
  });

  test("a 500 reply is unavailable", async () => {
    const error = await failureOf(
      createJsonApiConnector({
        url,
        columns,
        fetch: fakeFetch(reply("oops", { status: 500 })),
      }).read(),
    );
    expectSourceError(error, "unavailable");
  });

  test("a reply that is not JSON is invalid", async () => {
    const error = await failureOf(
      createJsonApiConnector({
        url,
        columns,
        fetch: fakeFetch(reply("<html>not json</html>")),
      }).read(),
    );
    expectSourceError(error, "invalid");
  });

  test("a reply where the rows path is not a list is invalid", async () => {
    const error = await failureOf(
      createJsonApiConnector({
        url,
        rowsPath: "data",
        columns,
        fetch: fakeFetch(reply({ data: { item: "widget" } })),
      }).read(),
    );
    expectSourceError(error, "invalid");
  });

  test("a content-length header over the limit is refused with code limit", async () => {
    const error = await failureOf(
      createJsonApiConnector({
        url,
        columns,
        fetch: fakeFetch(
          reply("[]", {
            headers: {
              "content-length": String(externalSourceLimits.responseBytes + 1),
            },
          }),
        ),
      }).read(),
    );
    expectSourceError(error, "limit");
  });

  test("a rejected fetch is unavailable, and an AbortError is cancelled", async () => {
    const rejected = await failureOf(
      createJsonApiConnector({
        url,
        columns,
        fetch: fakeFetch(new TypeError("network down")),
      }).read(),
    );
    expectSourceError(rejected, "unavailable");
    const aborted = new Error("aborted");
    aborted.name = "AbortError";
    const cancelled = await failureOf(
      createJsonApiConnector({
        url,
        columns,
        fetch: fakeFetch(aborted),
      }).read(),
    );
    expectSourceError(cancelled, "cancelled");
  });

  test("a value of the wrong type is a schema error", async () => {
    const error = await failureOf(
      createJsonApiConnector({
        url,
        columns,
        fetch: fakeFetch(reply([{ item: "widget", qty: "3" }])),
      }).read(),
    );
    expectSourceError(error, "schema");
  });

  test("construction throws TypeError for a relative URL or a file: URL", () => {
    expect(() =>
      createJsonApiConnector({ url: "/relative/items", columns }),
    ).toThrow(TypeError);
    expect(() =>
      createJsonApiConnector({ url: "file:///etc/passwd", columns }),
    ).toThrow(TypeError);
  });
});

describe("createODataConnector", () => {
  const columns = [
    { name: "Id", type: "text" },
    { name: "Amount", type: "number" },
  ];
  const url = "https://odata.example.test/odata/Invoices";

  test("asks for the declared fields in $select and applies a fixed $filter", async () => {
    const fetch = fakeFetch(reply({ value: [] }));
    await createODataConnector({
      url,
      filter: "Amount gt 0",
      columns,
      fetch,
    }).read();
    const sent = new URL(fetch.calls[0].url);
    expect(sent.origin + sent.pathname).toBe(url);
    expect(sent.searchParams.get("$select")).toBe("Id,Amount");
    expect(sent.searchParams.get("$filter")).toBe("Amount gt 0");
  });

  test("sends no $filter when none is given", async () => {
    const fetch = fakeFetch(reply({ value: [] }));
    await createODataConnector({ url, columns, fetch }).read();
    const sent = new URL(fetch.calls[0].url);
    expect(sent.searchParams.get("$select")).toBe("Id,Amount");
    expect(sent.searchParams.has("$filter")).toBe(false);
  });

  test("follows absolute and relative @odata.nextLink pages and joins them", async () => {
    const fetch = fakeFetch(
      reply({
        value: [{ Id: "a", Amount: 1 }],
        "@odata.nextLink": `${url}?$skiptoken=2`,
      }),
      reply({
        value: [{ Id: "b", Amount: 2 }],
        "@odata.nextLink": "/odata/Invoices?$skiptoken=3",
      }),
      reply({ value: [{ Id: "c", Amount: 3 }] }),
    );
    const result = await createODataConnector({ url, columns, fetch }).read();
    expect(result.rows).toEqual([
      ["a", 1],
      ["b", 2],
      ["c", 3],
    ]);
    expect(fetch.calls.map((call) => call.url)).toEqual([
      "https://odata.example.test/odata/Invoices?%24select=Id%2CAmount",
      `${url}?$skiptoken=2`,
      "https://odata.example.test/odata/Invoices?$skiptoken=3",
    ]);
  });

  test("a nextLink on a different origin is refused and never fetched", async () => {
    const fetch = fakeFetch(
      reply({
        value: [],
        "@odata.nextLink": "https://elsewhere.example.test/steal?$skiptoken=2",
      }),
    );
    const error = await failureOf(
      createODataConnector({ url, columns, fetch }).read(),
    );
    expectSourceError(error, "refused");
    expect(fetch.calls).toHaveLength(1);
    expect(fetch.calls.some((call) => call.url.includes("elsewhere"))).toBe(
      false,
    );
  });

  test("a reply without a value list is invalid", async () => {
    const error = await failureOf(
      createODataConnector({
        url,
        columns,
        fetch: fakeFetch(reply({ other: 1 })),
      }).read(),
    );
    expectSourceError(error, "invalid");
  });

  test("rows over the limit across pages are refused with code limit", async () => {
    const half = Array.from({ length: 3000 }, (_, index) => ({
      Id: "x",
      Amount: index,
    }));
    const fetch = fakeFetch(
      reply({ value: half, "@odata.nextLink": `${url}?$skiptoken=2` }),
      reply({ value: half }),
    );
    const error = await failureOf(
      createODataConnector({ url, columns, fetch }).read(),
    );
    expectSourceError(error, "limit");
  });

  test("kind is odata", () => {
    expect(
      createODataConnector({ url, columns, fetch: fakeFetch() }).kind,
    ).toBe("odata");
  });
});

describe("createQueryConnector", () => {
  const columns = [
    { name: "item", type: "text" },
    { name: "qty", type: "number" },
  ];

  test("rows may be arrays in column order or objects keyed by column name", async () => {
    const asArrays = createQueryConnector({
      columns,
      run: () => [["widget", 3]],
    });
    const asObjects = createQueryConnector({
      columns,
      run: () => [{ qty: 3, item: "widget" }],
    });
    expect((await asArrays.read()).rows).toEqual([["widget", 3]]);
    expect((await asObjects.read()).rows).toEqual([["widget", 3]]);
  });

  test("kind defaults to database, and custom kinds are kept", () => {
    const run = () => [];
    expect(createQueryConnector({ columns, run }).kind).toBe("database");
    expect(createQueryConnector({ columns, run, kind: "odbc" }).kind).toBe(
      "odbc",
    );
    expect(createQueryConnector({ columns, run, kind: "erp" }).kind).toBe(
      "erp",
    );
  });

  test("run receives an object holding the read signal", async () => {
    let received;
    const controller = new AbortController();
    await createQueryConnector({
      columns,
      run: (context) => {
        received = context;
        return [];
      },
    }).read({ signal: controller.signal });
    expect(received.signal).toBe(controller.signal);
  });

  test("a run that throws is unavailable", async () => {
    const error = await failureOf(
      createQueryConnector({
        columns,
        run: () => {
          throw new Error("driver crashed");
        },
      }).read(),
    );
    expectSourceError(error, "unavailable");
  });

  test("values of the wrong type are a schema error", async () => {
    const error = await failureOf(
      createQueryConnector({
        columns,
        run: () => [["widget", "three"]],
      }).read(),
    );
    expectSourceError(error, "schema");
  });

  test("construction throws TypeError without a run function", () => {
    expect(() => createQueryConnector({ columns })).toThrow(TypeError);
  });
});

describe("createSqliteConnector", () => {
  let folder;
  let database;
  beforeAll(async () => {
    folder = await makeFolder();
    database = path.join(folder, "ledger.db");
    const handle = new DatabaseSync(database);
    handle.exec(
      "CREATE TABLE items (name TEXT, qty INTEGER, price REAL, note TEXT, placed TEXT)",
    );
    handle.exec(
      "INSERT INTO items VALUES ('bolt', 2, 1.5, NULL, '2026-03-04'), ('nut', 3, 2.25, 'bulk', '2026-03-05')",
    );
    handle.close();
  });
  afterAll(async () => {
    await fs.rm(folder, { recursive: true, force: true });
  });

  const columns = [
    { name: "name", type: "text" },
    { name: "qty", type: "number" },
    { name: "price", type: "number" },
    { name: "note", type: "text", nullable: true },
    { name: "placed", type: "date" },
  ];
  const select =
    "SELECT name, qty, price, note, placed FROM items ORDER BY name";

  test("reads a SELECT into typed rows with integer and real values as numbers", async () => {
    const built = createSqliteConnector({
      id: "items-db",
      label: "Items",
      database,
      query: select,
      columns,
    });
    expect(built.kind).toBe("database");
    const result = await built.read();
    expect(result.rows).toEqual([
      ["bolt", 2, 1.5, null, "2026-03-04"],
      ["nut", 3, 2.25, "bulk", "2026-03-05"],
    ]);
  });

  test("NULL is null in a nullable column and a schema error in a non-nullable one", async () => {
    const strict = createSqliteConnector({
      id: "items-db",
      label: "Items",
      database,
      query: select,
      columns: columns.map((column) => ({ ...column, nullable: false })),
    });
    expectSourceError(await failureOf(strict.read()), "schema");
  });

  test("parameters fill the ? placeholders", async () => {
    const result = await createSqliteConnector({
      id: "items-db",
      label: "Items",
      database,
      query:
        "SELECT name, qty, price, note, placed FROM items WHERE qty > ? ORDER BY name",
      parameters: [2],
      columns,
    }).read();
    expect(result.rows).toEqual([["nut", 3, 2.25, "bulk", "2026-03-05"]]);
  });

  test("a WITH ... SELECT query is accepted", async () => {
    const result = await createSqliteConnector({
      id: "items-db",
      label: "Items",
      database,
      query: `WITH base AS (SELECT * FROM items) ${select.replace("FROM items", "FROM base")}`,
      columns,
    }).read();
    expect(result.rows).toHaveLength(2);
  });

  test("a statement that is not a single SELECT is refused at construction", () => {
    for (const query of [
      "update items set qty = 0",
      "delete from items",
      "SELECT 1 FROM items; DROP TABLE items",
    ]) {
      expect(() =>
        createSqliteConnector({
          id: "items-db",
          label: "Items",
          database,
          query,
          columns,
        }),
      ).toThrow(TypeError);
    }
  });

  test("reading leaves the database file byte-for-byte unchanged", async () => {
    const before = await fs.readFile(database);
    await createSqliteConnector({
      id: "items-db",
      label: "Items",
      database,
      query: select,
      columns,
    }).read();
    const after = await fs.readFile(database);
    expect(after.equals(before)).toBe(true);
  });

  test("a missing database file is an ExternalSourceError with code unavailable", async () => {
    const error = await failureOf(
      createSqliteConnector({
        id: "items-db",
        label: "Items",
        database: path.join(folder, "missing.db"),
        query: select,
        columns,
      }).read(),
    );
    expectSourceError(error, "unavailable");
  });

  test("a query naming a missing table is an ExternalSourceError with code unavailable", async () => {
    const error = await failureOf(
      createSqliteConnector({
        id: "items-db",
        label: "Items",
        database,
        query: "SELECT name FROM no_such_table",
        columns: [{ name: "name", type: "text" }],
      }).read(),
    );
    expectSourceError(error, "unavailable");
  });
});

describe("source gateway over these connectors", () => {
  let folder;
  let database;
  beforeAll(async () => {
    folder = await makeFolder();
    await fs.writeFile(
      path.join(folder, "orders.csv"),
      "item,qty\nwidget,3\ngadget,5\n",
    );
    database = path.join(folder, "orders.db");
    const handle = new DatabaseSync(database);
    handle.exec("CREATE TABLE orders (item TEXT, qty INTEGER)");
    handle.exec("INSERT INTO orders VALUES ('bolt', 7)");
    handle.close();
  });
  afterAll(async () => {
    await fs.rm(folder, { recursive: true, force: true });
  });

  const columns = [
    { name: "item", type: "text" },
    { name: "qty", type: "number" },
  ];

  test("refresh returns the file connector's rows in the ledgerglass-sources-v1 protocol", async () => {
    const gateway = createSourceGateway({
      isAllowed: () => true,
      connectors: [
        createDelimitedFileConnector({
          id: "orders-file",
          label: "Orders file",
          directory: folder,
          file: "orders.csv",
          columns,
        }),
      ],
    });
    const payload = await gateway.refresh("orders-file");
    expect(payload.protocol).toBe("ledgerglass-sources-v1");
    expect(payload.source.id).toBe("orders-file");
    expect(payload.rows).toEqual([
      ["widget", 3],
      ["gadget", 5],
    ]);
  });

  test("refresh returns the SQLite connector's rows in the ledgerglass-sources-v1 protocol", async () => {
    const gateway = createSourceGateway({
      isAllowed: () => true,
      connectors: [
        createSqliteConnector({
          id: "orders-db",
          label: "Orders database",
          database,
          query: "SELECT item, qty FROM orders",
          columns,
        }),
      ],
    });
    const payload = await gateway.refresh("orders-db");
    expect(payload.protocol).toBe("ledgerglass-sources-v1");
    expect(payload.source.id).toBe("orders-db");
    expect(payload.rows).toEqual([["bolt", 7]]);
  });
});

describe("createSqliteConnector query check", () => {
  const base = {
    id: "q",
    label: "Q",
    database: "/nonexistent/never-opened.db",
    columns: [{ name: "v", type: "number" }],
  };
  const build = (query) => () => createSqliteConnector({ ...base, query });

  test("a WITH clause that ends in a write is refused when the connector is created", () => {
    expect(build("WITH x AS (SELECT 1 AS v) DELETE FROM items")).toThrow(
      TypeError,
    );
    expect(build("WITH x AS (SELECT 1) INSERT INTO t SELECT * FROM x")).toThrow(
      TypeError,
    );
  });

  test("a semicolon or a keyword inside a string literal is not mistaken for a statement", () => {
    expect(build("SELECT 1 AS v WHERE 'a;b' = 'a;b'")).not.toThrow();
    expect(build("SELECT 1 AS v WHERE 'delete' <> 'update'")).not.toThrow();
    expect(build('SELECT 1 AS "drop"')).not.toThrow();
  });

  test("a second statement hidden after a comment is refused", () => {
    expect(build("SELECT 1 AS v; /* x */ DROP TABLE t")).toThrow(TypeError);
    expect(build("SELECT 1 AS v -- note\n; DELETE FROM t")).toThrow(TypeError);
  });

  test("one trailing semicolon is allowed", () => {
    expect(build("SELECT 1 AS v;")).not.toThrow();
  });
});
