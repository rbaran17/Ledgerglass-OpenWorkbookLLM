import { once } from "node:events";
import { createServer } from "node:http";
import { afterAll, describe, expect, test, vi } from "vitest";
import {
  ExternalSourceError,
  createSourceGateway,
  createSourceGatewayHttpHandler,
  externalSourceLimits,
  isSourceId,
  parseSourceRequest,
  shapeSourceRows,
  sourceErrorStatus,
  toExternalSourceError,
} from "../../src/sources/gateway.js";
import { createHttpSourceAdapter } from "../../src/integrations/http-source-adapter.js";
import {
  authorizeSourceAccess,
  fetchSourceCatalog,
  refreshSourceSnapshot,
} from "../../src/app/54-workbook-save-conflicts.js";

const FIXED_NOW = Date.parse("2026-11-12T10:00:00.000Z");
const fixedClock = () => FIXED_NOW;

/** The error a synchronous call throws, or null. */
function thrownBy(fn) {
  try {
    fn();
  } catch (error) {
    return error;
  }
  return null;
}

/** The error an async call rejects with, or null. */
async function rejectionOf(promise) {
  try {
    await promise;
  } catch (error) {
    return error;
  }
  return null;
}

/** The protocol code a synchronous call throws, or a description of what it threw. */
function codeOf(fn) {
  const error = thrownBy(fn);
  if (error === null) return "no error";
  return error instanceof ExternalSourceError
    ? error.code
    : `not an ExternalSourceError: ${error.name}`;
}

function ledgerConnector(overrides = {}) {
  return {
    id: "ledger",
    label: "General ledger",
    kind: "database",
    columns: [
      { name: "Account", type: "text" },
      { name: "Amount", type: "number" },
    ],
    controlFields: ["Amount"],
    read: async () => ({
      rows: [
        ["Cash", 1200.5],
        ["Rent", -800],
      ],
    }),
    ...overrides,
  };
}

function makeGateway(options = {}, connectors = [ledgerConnector()]) {
  return createSourceGateway({
    connectors,
    isAllowed: () => true,
    now: fixedClock,
    ...options,
  });
}

/** Connectors for the HTTP tests: one allowed, one a caller may not see, and three that fail. */
function serverConnectors() {
  return [
    ledgerConnector(),
    ledgerConnector({ id: "secret", label: "Payroll" }),
    ledgerConnector({
      id: "typed",
      columns: [{ name: "Amount", type: "number" }],
      controlFields: ["Amount"],
      read: async () => ({ rows: [["12"]] }),
    }),
    ledgerConnector({
      id: "wide",
      columns: [{ name: "Amount", type: "number" }],
      controlFields: ["Amount"],
      read: async () => ({
        rows: Array.from({ length: externalSourceLimits.rows + 1 }, (_, i) => [
          i,
        ]),
      }),
    }),
    ledgerConnector({
      id: "broken",
      read: async () => {
        throw new Error("database is down");
      },
    }),
  ];
}

const servers = [];
afterAll(async () => {
  for (const server of servers) {
    server.closeAllConnections();
    await new Promise((resolve) => server.close(() => resolve()));
  }
});

/**
 * A real node:http server running the source endpoint on port 0. The caller is "alice" when the
 * x-user header says so; "alice" may see every source except "secret".
 */
async function startSourceServer({
  connectors = serverConnectors(),
  gateway = {},
  handler = {},
} = {}) {
  const sourceGateway = createSourceGateway({
    connectors,
    isAllowed: (sourceId, principal) =>
      principal?.name === "alice" && sourceId !== "secret",
    now: () => Date.now(),
    ...gateway,
  });
  const seen = [];
  const sourceHandler = createSourceGatewayHttpHandler(sourceGateway, {
    authenticate: (request) => {
      seen.push(request.headers);
      if (request.headers["x-user"] === "throw")
        throw new Error("identity provider is down");
      return request.headers["x-user"]
        ? { name: request.headers["x-user"] }
        : null;
    },
    ...handler,
  });
  const handled = [];
  const server = createServer(async (request, response) => {
    const claimed = await sourceHandler(request, response);
    handled.push(claimed);
    if (!claimed) {
      response.writeHead(404);
      response.end();
    }
  });
  server.listen(0, "127.0.0.1");
  await once(server, "listening");
  servers.push(server);
  const origin = `http://127.0.0.1:${server.address().port}`;
  return { origin, url: `${origin}/api/sources`, seen, handled, server };
}

function post(url, body, headers = {}) {
  return fetch(url, {
    method: "POST",
    headers: {
      "content-type": "application/json",
      "x-user": "alice",
      ...headers,
    },
    body: typeof body === "string" ? body : JSON.stringify(body),
  });
}

describe("protocol: parseSourceRequest", () => {
  test("accepts the three request shapes", () => {
    expect(parseSourceRequest({ operation: "catalog" })).toEqual({
      operation: "catalog",
    });
    expect(
      parseSourceRequest({ operation: "authorize", sourceId: "ledger" }),
    ).toEqual({ operation: "authorize", sourceId: "ledger" });
    expect(
      parseSourceRequest({ operation: "refresh", sourceId: "ledger" }),
    ).toEqual({ operation: "refresh", sourceId: "ledger" });
  });

  test("rejects extra keys", () => {
    expect(
      codeOf(() =>
        parseSourceRequest({ operation: "catalog", sourceId: "ledger" }),
      ),
    ).toBe("invalid");
    expect(
      codeOf(() =>
        parseSourceRequest({
          operation: "refresh",
          sourceId: "ledger",
          extra: true,
        }),
      ),
    ).toBe("invalid");
  });

  test("rejects a missing or invalid sourceId", () => {
    expect(codeOf(() => parseSourceRequest({ operation: "refresh" }))).toBe(
      "invalid",
    );
    expect(
      codeOf(() => parseSourceRequest({ operation: "refresh", sourceId: 42 })),
    ).toBe("invalid");
    expect(
      codeOf(() =>
        parseSourceRequest({ operation: "authorize", sourceId: "bad id" }),
      ),
    ).toBe("invalid");
  });

  test("rejects unknown operations", () => {
    for (const body of [
      { operation: "delete", sourceId: "ledger" },
      { operation: "Catalog" },
      { operation: "" },
      {},
    ])
      expect(codeOf(() => parseSourceRequest(body))).toBe("invalid");
  });

  test("rejects non-object bodies", () => {
    for (const body of [null, undefined, "catalog", 42, true, [], ["catalog"]])
      expect(codeOf(() => parseSourceRequest(body))).toBe("invalid");
  });
});

describe("protocol: isSourceId", () => {
  test("accepts 1 to 100 identifier characters and rejects the edges", () => {
    expect(isSourceId("a")).toBe(true);
    expect(isSourceId("a".repeat(100))).toBe(true);
    expect(isSourceId("a".repeat(101))).toBe(false);
    expect(isSourceId("a.b-c_d9")).toBe(true);
    expect(isSourceId(".ledger")).toBe(false);
    expect(isSourceId("-ledger")).toBe(false);
    expect(isSourceId("general ledger")).toBe(false);
    expect(isSourceId("")).toBe(false);
    expect(isSourceId("abc\n")).toBe(false);
    expect(isSourceId(7)).toBe(false);
  });
});

describe("protocol: toExternalSourceError", () => {
  test("keeps an ExternalSourceError as it is", () => {
    const error = new ExternalSourceError("refused");
    expect(toExternalSourceError(error)).toBe(error);
  });

  test("keeps the code of a look-alike error only when the code is known", () => {
    const known = toExternalSourceError(
      Object.assign(new Error("x"), {
        name: "ExternalSourceError",
        code: "refused",
      }),
    );
    expect(known).toBeInstanceOf(ExternalSourceError);
    expect(known.code).toBe("refused");
    expect(
      toExternalSourceError(
        Object.assign(new Error("x"), {
          name: "ExternalSourceError",
          code: "melted",
        }),
      ).code,
    ).toBe("unavailable");
  });

  test("maps AbortError to cancelled and TimeoutError to timeout", () => {
    expect(
      toExternalSourceError(
        Object.assign(new Error("a"), { name: "AbortError" }),
      ).code,
    ).toBe("cancelled");
    expect(
      toExternalSourceError(
        Object.assign(new Error("t"), { name: "TimeoutError" }),
      ).code,
    ).toBe("timeout");
  });

  test("maps anything else to unavailable", () => {
    for (const error of [new Error("boom"), "boom", null, undefined])
      expect(toExternalSourceError(error).code).toBe("unavailable");
  });
});

describe("protocol: shapeSourceRows", () => {
  const columns = [
    { name: "Account", type: "text" },
    { name: "Amount", type: "number" },
  ];

  test("reads arrays in column order and keyed objects by column name", () => {
    expect(shapeSourceRows([["Cash", 1200.5]], columns)).toEqual([
      ["Cash", 1200.5],
    ]);
    expect(shapeSourceRows([{ Amount: 3, Account: "Rent" }], columns)).toEqual([
      ["Rent", 3],
    ]);
  });

  test("number columns take finite numbers and safe bigints only", () => {
    const numeric = [{ name: "n", type: "number" }];
    expect(shapeSourceRows([[7n]], numeric)).toEqual([[7]]);
    expect(codeOf(() => shapeSourceRows([[NaN]], numeric))).toBe("schema");
    expect(codeOf(() => shapeSourceRows([[Infinity]], numeric))).toBe("schema");
    expect(codeOf(() => shapeSourceRows([[2n ** 60n]], numeric))).toBe(
      "schema",
    );
    expect(codeOf(() => shapeSourceRows([["12"]], numeric))).toBe("schema");
  });

  test("text and boolean columns refuse other types; text is kept as written", () => {
    expect(codeOf(() => shapeSourceRows([[5, 1]], columns))).toBe("schema");
    const flags = [{ name: "flag", type: "boolean" }];
    expect(codeOf(() => shapeSourceRows([["true"]], flags))).toBe("schema");
    expect(codeOf(() => shapeSourceRows([[1]], flags))).toBe("schema");
    expect(shapeSourceRows([[""]], [{ name: "t", type: "text" }])).toEqual([
      [""],
    ]);
  });

  test("with parseText, text numbers are read as numbers; otherwise they are refused", () => {
    const numeric = [{ name: "n", type: "number" }];
    expect(shapeSourceRows([["1200.5"]], numeric, { parseText: true })).toEqual(
      [[1200.5]],
    );
    expect(
      shapeSourceRows([["  -1e3 "]], numeric, { parseText: true }),
    ).toEqual([[-1000]]);
    expect(shapeSourceRows([[".5"]], numeric, { parseText: true })).toEqual([
      [0.5],
    ]);
    expect(
      codeOf(() => shapeSourceRows([["1,200"]], numeric, { parseText: true })),
    ).toBe("schema");
    expect(codeOf(() => shapeSourceRows([["1200.5"]], numeric))).toBe("schema");
  });

  test("with parseText, an empty string is null for a nullable column and refused otherwise", () => {
    const nullableNumber = [{ name: "n", type: "number", nullable: true }];
    const requiredNumber = [{ name: "n", type: "number" }];
    expect(
      shapeSourceRows([[""]], nullableNumber, { parseText: true }),
    ).toEqual([[null]]);
    expect(
      codeOf(() =>
        shapeSourceRows([[""]], requiredNumber, { parseText: true }),
      ),
    ).toBe("schema");
    expect(codeOf(() => shapeSourceRows([[""]], requiredNumber))).toBe(
      "schema",
    );
  });

  test("with parseText, booleans are read case-insensitively; without it, strings are refused", () => {
    const flags = [{ name: "flag", type: "boolean" }];
    expect(
      shapeSourceRows([["TRUE"], [" false "]], flags, { parseText: true }),
    ).toEqual([[true], [false]]);
    expect(
      codeOf(() => shapeSourceRows([["yes"]], flags, { parseText: true })),
    ).toBe("schema");
  });

  test("date columns accept Date objects and date-times, and refuse impossible dates", () => {
    const dates = [{ name: "d", type: "date" }];
    expect(
      shapeSourceRows(
        [
          [new Date("2026-03-01T12:00:00Z")],
          ["2026-03-01T23:59:59Z"],
          ["2026-03-01"],
        ],
        dates,
      ),
    ).toEqual([["2026-03-01"], ["2026-03-01"], ["2026-03-01"]]);
    expect(codeOf(() => shapeSourceRows([["2026-02-30"]], dates))).toBe(
      "schema",
    );
    expect(
      codeOf(() => shapeSourceRows([["2026-02-30T00:00:00Z"]], dates)),
    ).toBe("schema");
    expect(codeOf(() => shapeSourceRows([["2026/03/01"]], dates))).toBe(
      "schema",
    );
    expect(codeOf(() => shapeSourceRows([[new Date("nope")]], dates))).toBe(
      "schema",
    );
  });

  test("null and missing values are allowed only in nullable columns", () => {
    const nullable = [{ name: "n", type: "number", nullable: true }];
    const required = [{ name: "n", type: "number" }];
    expect(shapeSourceRows([[null], [undefined]], nullable)).toEqual([
      [null],
      [null],
    ]);
    expect(codeOf(() => shapeSourceRows([[null]], required))).toBe("schema");
    expect(codeOf(() => shapeSourceRows([{}], required))).toBe("schema");
  });

  test("row and cell limits are enforced at the published boundaries", () => {
    const one = [{ name: "n", type: "number" }];
    const tooManyRows = Array.from(
      { length: externalSourceLimits.rows + 1 },
      () => [1],
    );
    expect(codeOf(() => shapeSourceRows(tooManyRows, one))).toBe("limit");
    const wide = Array.from({ length: 50 }, (_, i) => ({
      name: `c${i}`,
      type: "number",
      nullable: true,
    }));
    const cellsOk = Array.from({ length: 1999 }, () => Array(50).fill(0));
    const cellsOver = Array.from({ length: 2000 }, () => Array(50).fill(0));
    expect(shapeSourceRows(cellsOk, wide)).toHaveLength(1999);
    expect(codeOf(() => shapeSourceRows(cellsOver, wide))).toBe("limit");
  });

  test("non-array input is invalid; a text cell over 32767 characters is a limit", () => {
    expect(codeOf(() => shapeSourceRows({ a: 1 }, columns))).toBe("invalid");
    const text = [{ name: "t", type: "text" }];
    expect(shapeSourceRows([["x".repeat(32767)]], text)).toHaveLength(1);
    expect(codeOf(() => shapeSourceRows([["x".repeat(32768)]], text))).toBe(
      "limit",
    );
  });
});

describe("createSourceGateway construction", () => {
  const build = (connectors, options = {}) =>
    createSourceGateway({ connectors, isAllowed: () => true, ...options });

  test("requires isAllowed, and between 1 and 20 connectors", () => {
    expect(() =>
      createSourceGateway({ connectors: [ledgerConnector()] }),
    ).toThrow(TypeError);
    expect(() =>
      createSourceGateway({
        connectors: [ledgerConnector()],
        isAllowed: "yes",
      }),
    ).toThrow(/isAllowed/);
    expect(() => build([])).toThrow(TypeError);
    const tooMany = Array.from(
      { length: externalSourceLimits.sources + 1 },
      (_, i) => ledgerConnector({ id: `c${i}` }),
    );
    expect(() => build(tooMany)).toThrow(/1 to 20/);
  });

  test("rejects a connector with a bad id, a bad kind, or no read function", () => {
    expect(() => build([ledgerConnector({ id: ".hidden" })])).toThrow(/id/);
    expect(() => build([ledgerConnector({ kind: "spreadsheet" })])).toThrow(
      /kind/,
    );
    expect(() => build([ledgerConnector({ read: undefined })])).toThrow(/read/);
  });

  test("rejects a bad column type and duplicate column names, ignoring case", () => {
    expect(() =>
      build([
        ledgerConnector({
          columns: [{ name: "Account", type: "string" }],
          controlFields: [],
        }),
      ]),
    ).toThrow(/type/);
    expect(() =>
      build([
        ledgerConnector({
          columns: [
            { name: "Amount", type: "number" },
            { name: "amount", type: "number" },
          ],
        }),
      ]),
    ).toThrow(/listed twice/);
  });

  test("rejects a control field that is not a number column", () => {
    expect(() =>
      build([ledgerConnector({ controlFields: ["Account"] })]),
    ).toThrow(/control field "Account"/);
    expect(() =>
      build([ledgerConnector({ controlFields: ["Missing"] })]),
    ).toThrow(/control field "Missing"/);
  });

  test("rejects two connectors with the same id", () => {
    expect(() => build([ledgerConnector(), ledgerConnector()])).toThrow(
      /registered twice/,
    );
  });

  test("approvalTtlMs must be between one second and 24 hours", () => {
    for (const approvalTtlMs of [999, 1.5, 24 * 60 * 60 * 1000 + 1])
      expect(() => build([ledgerConnector()], { approvalTtlMs })).toThrow(
        /approvalTtlMs/,
      );
    expect(() =>
      build([ledgerConnector()], { approvalTtlMs: 1000 }),
    ).not.toThrow();
    expect(() =>
      build([ledgerConnector()], { approvalTtlMs: 24 * 60 * 60 * 1000 }),
    ).not.toThrow();
  });
});

describe("gateway: catalog, authorize and refresh", () => {
  test("catalog lists only allowed sources, each with expiresAt = now plus the ttl", async () => {
    const gateway = makeGateway(
      {
        isAllowed: (sourceId) => sourceId === "ledger",
        approvalTtlMs: 60_000,
      },
      [ledgerConnector(), ledgerConnector({ id: "payroll" })],
    );
    const approvals = await gateway.catalog({ principal: "alice" });
    expect(approvals).toHaveLength(1);
    expect(approvals[0].source.id).toBe("ledger");
    expect(approvals[0].expiresAt).toBe("2026-11-12T10:01:00.000Z");
    expect(Object.keys(approvals[0]).sort()).toEqual(["expiresAt", "source"]);
  });

  test("isAllowed receives the source id and the caller's principal", async () => {
    const isAllowed = vi.fn(() => true);
    const gateway = makeGateway({ isAllowed });
    const principal = { name: "alice" };
    await gateway.authorize("ledger", { principal });
    expect(isAllowed).toHaveBeenCalledWith("ledger", principal);
  });

  test("an isAllowed that throws, or rejects, counts as not allowed", async () => {
    const throwing = makeGateway({
      isAllowed: () => {
        throw new Error("directory is down");
      },
    });
    expect(await throwing.catalog({ principal: "alice" })).toEqual([]);
    const error = await rejectionOf(
      throwing.authorize("ledger", { principal: "alice" }),
    );
    expect(error.code).toBe("refused");
    const rejecting = makeGateway({
      isAllowed: async () => {
        throw new Error("directory is down");
      },
    });
    expect((await rejectionOf(rejecting.refresh("ledger", {}))).code).toBe(
      "refused",
    );
  });

  test("isAllowed must answer exactly true; another truthy value is refused", async () => {
    const gateway = makeGateway({ isAllowed: () => 1 });
    expect(await gateway.catalog({ principal: "alice" })).toEqual([]);
    expect((await rejectionOf(gateway.authorize("ledger", {}))).code).toBe(
      "refused",
    );
  });

  test("an unknown source and a disallowed source are refused the same way", async () => {
    const gateway = makeGateway(
      { isAllowed: (sourceId) => sourceId !== "secret" },
      [ledgerConnector(), ledgerConnector({ id: "secret" })],
    );
    const unknown = await rejectionOf(gateway.authorize("nope", {}));
    const disallowed = await rejectionOf(gateway.authorize("secret", {}));
    expect(unknown.code).toBe("refused");
    expect(disallowed.code).toBe("refused");
    expect(unknown.message).toBe(disallowed.message);
  });

  test("schemaVersion is stable for the same columns and changes when a column changes", async () => {
    const versionOf = async (connector) =>
      (await makeGateway({}, [connector]).authorize(connector.id, {})).source
        .schemaVersion;
    const first = await versionOf(ledgerConnector());
    expect(await versionOf(ledgerConnector())).toBe(first);
    expect(
      await versionOf(
        ledgerConnector({
          columns: [
            { name: "Account", type: "text" },
            { name: "Amount", type: "text" },
          ],
          controlFields: [],
        }),
      ),
    ).not.toBe(first);
  });

  test("refresh returns the full protocol payload for the approved source", async () => {
    const payload = await makeGateway().refresh("ledger", {
      principal: "alice",
    });
    expect(payload).toEqual({
      protocol: "ledgerglass-sources-v1",
      source: {
        id: "ledger",
        label: "General ledger",
        kind: "database",
        columns: [
          { name: "Account", type: "text", nullable: false },
          { name: "Amount", type: "number", nullable: false },
        ],
        controlFields: ["Amount"],
        schemaVersion: expect.stringMatching(/^s-/),
      },
      sourceTimestamp: "2026-11-12T10:00:00.000Z",
      version: expect.stringMatching(/^v-/),
      rows: [
        ["Cash", 1200.5],
        ["Rent", -800],
      ],
      complete: true,
    });
  });

  test("version is the same for the same rows and different for different rows", async () => {
    const versionFor = async (rows) =>
      (
        await makeGateway({}, [
          ledgerConnector({ read: async () => ({ rows }) }),
        ]).refresh("ledger", {})
      ).version;
    const rows = [
      ["Cash", 1200.5],
      ["Rent", -800],
    ];
    expect(await versionFor(rows)).toBe(await versionFor([...rows]));
    expect(await versionFor(rows)).not.toBe(
      await versionFor([
        ["Cash", 1200.5],
        ["Rent", -801],
      ]),
    );
  });

  test("the connector's own version and sourceTimestamp are used when it gives them", async () => {
    const withStrings = makeGateway({}, [
      ledgerConnector({
        read: async () => ({
          rows: [["Cash", 1]],
          version: "ledger-2026-11-12",
          sourceTimestamp: "2026-11-12T09:00:00+02:00",
        }),
      }),
    ]);
    const first = await withStrings.refresh("ledger", {});
    expect(first.version).toBe("ledger-2026-11-12");
    expect(first.sourceTimestamp).toBe("2026-11-12T07:00:00.000Z");

    const withDate = makeGateway({}, [
      ledgerConnector({
        read: async () => ({
          rows: [["Cash", 1]],
          sourceTimestamp: new Date("2026-11-11T23:30:00Z"),
        }),
      }),
    ]);
    expect((await withDate.refresh("ledger", {})).sourceTimestamp).toBe(
      "2026-11-11T23:30:00.000Z",
    );
  });

  test("a connector's failure is unavailable, unless it is already a timeout", async () => {
    const plain = makeGateway({}, [
      ledgerConnector({
        read: async () => {
          throw new Error("socket hang up");
        },
      }),
    ]);
    expect((await rejectionOf(plain.refresh("ledger", {}))).code).toBe(
      "unavailable",
    );
    const timed = makeGateway({}, [
      ledgerConnector({
        read: async () => {
          throw new ExternalSourceError("timeout");
        },
      }),
    ]);
    expect((await rejectionOf(timed.refresh("ledger", {}))).code).toBe(
      "timeout",
    );
  });

  test("rows of the wrong width, the wrong type, or a null in a required column are schema errors", async () => {
    const badRows = [[["Cash"]], [["Cash", "1200"]], [["Cash", null]]];
    for (const rows of badRows) {
      const gateway = makeGateway({}, [
        ledgerConnector({ read: async () => ({ rows }) }),
      ]);
      expect((await rejectionOf(gateway.refresh("ledger", {}))).code).toBe(
        "schema",
      );
    }
  });

  test("too many rows is a limit error", async () => {
    const gateway = makeGateway({}, [
      ledgerConnector({
        read: async () => ({
          rows: Array.from({ length: externalSourceLimits.rows + 1 }, () => [
            "a",
            1,
          ]),
        }),
      }),
    ]);
    expect((await rejectionOf(gateway.refresh("ledger", {}))).code).toBe(
      "limit",
    );
  });

  test("an aborted request is cancelled before the connector runs; a live one passes its signal on", async () => {
    const read = vi.fn(async () => ({ rows: [["Cash", 1]] }));
    const gateway = makeGateway({}, [ledgerConnector({ read })]);
    const aborted = await rejectionOf(
      gateway.refresh("ledger", {
        principal: "alice",
        signal: AbortSignal.abort(),
      }),
    );
    expect(aborted.code).toBe("cancelled");
    expect(read).not.toHaveBeenCalled();

    const controller = new AbortController();
    await gateway.refresh("ledger", {
      principal: "alice",
      signal: controller.signal,
    });
    expect(read.mock.calls[0][0].signal).toBe(controller.signal);
  });

  test("handle dispatches all three operations and refuses a bad body as invalid", async () => {
    const gateway = makeGateway();
    const context = { principal: "alice" };
    expect(
      await gateway.handle({ operation: "catalog" }, context),
    ).toHaveLength(1);
    expect(
      (
        await gateway.handle(
          { operation: "authorize", sourceId: "ledger" },
          context,
        )
      ).source.id,
    ).toBe("ledger");
    expect(
      (
        await gateway.handle(
          { operation: "refresh", sourceId: "ledger" },
          context,
        )
      ).complete,
    ).toBe(true);
    expect(
      (await rejectionOf(gateway.handle({ operation: "delete" }, context)))
        .code,
    ).toBe("invalid");
    expect((await rejectionOf(gateway.handle(null, context))).code).toBe(
      "invalid",
    );
  });
});

describe("HTTP handler over a real server", () => {
  test("requires authenticate", () => {
    expect(() => createSourceGatewayHttpHandler(makeGateway(), {})).toThrow(
      /authenticate/,
    );
  });

  test("answers 404 for another path, and 405 with Allow: POST for a GET", async () => {
    const source = await startSourceServer();
    const elsewhere = await fetch(`${source.origin}/elsewhere`, {
      method: "POST",
      headers: { "content-type": "application/json", "x-user": "alice" },
      body: "{}",
    });
    expect(elsewhere.status).toBe(404);
    expect(source.handled.at(-1)).toBe(false);

    const wrongMethod = await fetch(source.url, {
      headers: { "x-user": "alice" },
    });
    expect(wrongMethod.status).toBe(405);
    expect(wrongMethod.headers.get("allow")).toBe("POST");
  });

  test("415 without a JSON content type", async () => {
    const source = await startSourceServer();
    const response = await post(source.url, '{"operation":"catalog"}', {
      "content-type": "text/plain",
    });
    expect(response.status).toBe(415);
    expect((await response.json()).error.code).toBe("invalid");
  });

  test("401 when authenticate returns nothing or throws", async () => {
    const source = await startSourceServer();
    const anonymous = await post(
      source.url,
      { operation: "catalog" },
      { "x-user": "" },
    );
    expect(anonymous.status).toBe(401);
    expect((await anonymous.json()).error.code).toBe("refused");
    const thrown = await post(
      source.url,
      { operation: "catalog" },
      { "x-user": "throw" },
    );
    expect(thrown.status).toBe(401);
  });

  test("200 with JSON for each operation, marked no-store", async () => {
    const source = await startSourceServer();
    const catalog = await post(source.url, { operation: "catalog" });
    expect(catalog.status).toBe(200);
    expect(catalog.headers.get("cache-control")).toBe("no-store");
    expect(catalog.headers.get("content-type")).toMatch(/^application\/json/);
    expect(
      (await catalog.json()).map((approval) => approval.source.id),
    ).toEqual(["ledger", "typed", "wide", "broken"]);
    const authorize = await post(source.url, {
      operation: "authorize",
      sourceId: "ledger",
    });
    expect(authorize.status).toBe(200);
    expect((await authorize.json()).source.id).toBe("ledger");
    const refresh = await post(source.url, {
      operation: "refresh",
      sourceId: "ledger",
    });
    expect(refresh.status).toBe(200);
    expect((await refresh.json()).rows).toEqual([
      ["Cash", 1200.5],
      ["Rent", -800],
    ]);
  });

  test("errors carry { error: { code, message } } with the status for each code", async () => {
    const source = await startSourceServer();
    const cases = [
      ["secret", 403, "refused"],
      ["nope", 403, "refused"],
      ["typed", 422, "schema"],
      ["wide", 413, "limit"],
      ["broken", 503, "unavailable"],
    ];
    for (const [sourceId, status, code] of cases) {
      const response = await post(source.url, {
        operation: "refresh",
        sourceId,
      });
      expect(response.status, sourceId).toBe(status);
      const body = await response.json();
      expect(Object.keys(body)).toEqual(["error"]);
      expect(Object.keys(body.error).sort()).toEqual(["code", "message"]);
      expect(body.error.code, sourceId).toBe(code);
    }
  });

  test("bad JSON and an oversized body are both 400 invalid", async () => {
    const source = await startSourceServer();
    const malformed = await post(source.url, "{not json");
    expect(malformed.status).toBe(400);
    expect((await malformed.json()).error.code).toBe("invalid");
    const oversized = await post(source.url, {
      operation: "catalog",
      pad: "x".repeat(5000),
    });
    expect(oversized.status).toBe(400);
    expect((await oversized.json()).error.code).toBe("invalid");
  });

  test("a connector slower than timeoutMs answers 504 timeout", async () => {
    const source = await startSourceServer({
      connectors: [
        ledgerConnector({
          id: "slow",
          read: ({ signal }) =>
            new Promise((_, reject) =>
              signal.addEventListener("abort", () => reject(signal.reason)),
            ),
        }),
      ],
      handler: { timeoutMs: 50 },
    });
    const response = await post(source.url, {
      operation: "refresh",
      sourceId: "slow",
    });
    expect(response.status).toBe(504);
    expect((await response.json()).error.code).toBe("timeout");
  });
});

describe("HTTP source adapter over a real server", () => {
  test("needs an endpoint", () => {
    expect(() => createHttpSourceAdapter({})).toThrow(/endpoint/);
  });

  test("catalog, authorize and refresh round trip through the server", async () => {
    const source = await startSourceServer();
    const adapter = createHttpSourceAdapter({
      endpoint: source.url,
      headers: { "x-user": "alice" },
    });
    const catalog = await adapter.catalog();
    expect(catalog.map((approval) => approval.source.id)).toEqual([
      "ledger",
      "typed",
      "wide",
      "broken",
    ]);
    expect((await adapter.authorize("ledger")).source.id).toBe("ledger");
    expect((await adapter.refresh("ledger")).rows).toHaveLength(2);
  });

  test("error codes arrive as ExternalSourceError with the server's code", async () => {
    const source = await startSourceServer();
    const adapter = createHttpSourceAdapter({
      endpoint: source.url,
      headers: { "x-user": "alice" },
    });
    for (const [sourceId, code] of [
      ["secret", "refused"],
      ["typed", "schema"],
      ["wide", "limit"],
      ["broken", "unavailable"],
    ]) {
      const error = await rejectionOf(adapter.refresh(sourceId));
      expect(error, sourceId).toBeInstanceOf(ExternalSourceError);
      expect(error.code, sourceId).toBe(code);
    }
  });

  test("a 401 from the server becomes refused", async () => {
    const source = await startSourceServer();
    const adapter = createHttpSourceAdapter({ endpoint: source.url });
    expect((await rejectionOf(adapter.catalog())).code).toBe("refused");
  });

  test("a server that is down is unavailable", async () => {
    const source = await startSourceServer();
    source.server.closeAllConnections();
    await new Promise((resolve) => source.server.close(() => resolve()));
    const adapter = createHttpSourceAdapter({
      endpoint: source.url,
      headers: { "x-user": "alice" },
    });
    expect((await rejectionOf(adapter.catalog())).code).toBe("unavailable");
  });

  test("an aborted signal is cancelled", async () => {
    const source = await startSourceServer();
    const adapter = createHttpSourceAdapter({
      endpoint: source.url,
      headers: { "x-user": "alice" },
    });
    const controller = new AbortController();
    controller.abort();
    expect(
      (await rejectionOf(adapter.catalog({ signal: controller.signal }))).code,
    ).toBe("cancelled");
  });

  test("custom headers are sent with every request", async () => {
    const source = await startSourceServer();
    const adapter = createHttpSourceAdapter({
      endpoint: source.url,
      headers: { "x-user": "alice", "x-trace": "trace-123" },
    });
    await adapter.catalog();
    expect(source.seen.at(-1)["x-trace"]).toBe("trace-123");
    expect(source.seen.at(-1)["content-type"]).toBe("application/json");
  });
});

describe("the recovered editor accepts the server", () => {
  test("fetchSourceCatalog and authorizeSourceAccess return the parsed approvals", async () => {
    const source = await startSourceServer();
    const adapter = createHttpSourceAdapter({
      endpoint: source.url,
      headers: { "x-user": "alice" },
    });
    const catalog = await fetchSourceCatalog(adapter);
    expect(catalog.map((approval) => approval.source.id)).toEqual([
      "ledger",
      "typed",
      "wide",
      "broken",
    ]);
    const grant = await authorizeSourceAccess(adapter, "ledger");
    expect(grant.source.id).toBe("ledger");
    expect(Date.parse(grant.expiresAt)).toBeGreaterThan(Date.now());
  });

  test("refreshSourceSnapshot returns the grant and the validated snapshot with rows", async () => {
    const source = await startSourceServer();
    const adapter = createHttpSourceAdapter({
      endpoint: source.url,
      headers: { "x-user": "alice" },
    });
    const { grant, snapshot } = await refreshSourceSnapshot(adapter, "ledger");
    expect(grant.source.id).toBe("ledger");
    expect(snapshot.protocol).toBe("ledgerglass-sources-v1");
    expect(snapshot.complete).toBe(true);
    expect(snapshot.rows).toEqual([
      ["Cash", 1200.5],
      ["Rent", -800],
    ]);
  });

  test("a disallowed source is rejected as refused", async () => {
    const source = await startSourceServer();
    const adapter = createHttpSourceAdapter({
      endpoint: source.url,
      headers: { "x-user": "alice" },
    });
    expect(
      (await rejectionOf(refreshSourceSnapshot(adapter, "secret"))).code,
    ).toBe("refused");
  });

  test("a connector returning wrong-typed rows is rejected as schema", async () => {
    const source = await startSourceServer();
    const adapter = createHttpSourceAdapter({
      endpoint: source.url,
      headers: { "x-user": "alice" },
    });
    expect(
      (await rejectionOf(refreshSourceSnapshot(adapter, "typed"))).code,
    ).toBe("schema");
  });

  test("an expired approval makes authorizeSourceAccess reject as expired", async () => {
    const source = await startSourceServer({
      gateway: { approvalTtlMs: 1000, now: () => Date.now() - 60_000 },
    });
    const adapter = createHttpSourceAdapter({
      endpoint: source.url,
      headers: { "x-user": "alice" },
    });
    expect(
      (await rejectionOf(authorizeSourceAccess(adapter, "ledger"))).code,
    ).toBe("expired");
  });
});

describe("fixes found by testing", () => {
  const column = (type, nullable = false) => [{ name: "v", type, nullable }];

  test("a number too large to represent is a schema error, not Infinity", () => {
    expect(() =>
      shapeSourceRows([["1e400"]], column("number"), { parseText: true }),
    ).toThrow(expect.objectContaining({ code: "schema" }));
  });

  test("in a text source, blank space stays text in a text column and is empty elsewhere", () => {
    expect(
      shapeSourceRows([["   "]], column("text"), { parseText: true }),
    ).toEqual([["   "]]);
    expect(
      shapeSourceRows([[""]], column("text", true), { parseText: true }),
    ).toEqual([[null]]);
    expect(
      shapeSourceRows([["  "]], column("number", true), { parseText: true }),
    ).toEqual([[null]]);
  });

  test("a source the editor would reject cannot be registered", () => {
    const connector = (overrides) => ({
      id: "s",
      label: "S",
      kind: "api",
      columns: [{ name: "amount", type: "number" }],
      read: async () => ({ rows: [] }),
      ...overrides,
    });
    const build = (overrides) => () =>
      createSourceGateway({
        isAllowed: () => true,
        connectors: [connector(overrides)],
      });
    expect(build({})).not.toThrow();
    expect(build({ controlFields: ["amount", "amount"] })).toThrow(TypeError);
    expect(build({ label: "x".repeat(129) })).toThrow(TypeError);
    expect(build({ label: "bad\u0007label" })).toThrow(TypeError);
    expect(build({ columns: [{ name: "a\tb", type: "text" }] })).toThrow(
      TypeError,
    );
  });
});
