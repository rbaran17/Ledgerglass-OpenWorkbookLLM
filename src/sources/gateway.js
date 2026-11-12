/**
 * Source gateway: the server side of the external source protocol. It sits between the editor
 * (which may only ask for a catalog, an approval, or a refresh of a named source) and the
 * connectors that reach the data. The browser never sends a query, a path or a URL: it names a
 * source the host registered, and the host decides who may see it.
 *
 * New code, written against the client contract recovered in src/sources/protocol.js.
 */
import { createHash } from "node:crypto";
import {
  ExternalSourceError,
  externalSourceFormatVersion,
  externalSourceKinds,
  externalSourceLimits,
  isSourceId,
  parseActiveSourceApproval,
  parseSourceRequest,
  toExternalSourceError,
  validateExternalSourcePayload,
} from "./protocol.js";

/** The protocol's names, limits and checks, for hosts that write their own connectors. */
export {
  ExternalSourceError,
  externalSourceErrorCodes,
  externalSourceFormatVersion,
  externalSourceKinds,
  externalSourceLimits,
  isSourceId,
  parseSourceRequest,
  shapeSourceRows,
  toExternalSourceError,
  validateExternalSourcePayload,
} from "./protocol.js";

const columnTypes = new Set(["text", "number", "boolean", "date"]);

function sha256Hex(text) {
  return createHash("sha256").update(text).digest("hex");
}

/** Checks a connector's description once, when the gateway is created, and returns the source table. */
function describeSource(connector) {
  const fail = (reason) => {
    throw new TypeError(
      `Source connector "${connector?.id ?? "?"}": ${reason}`,
    );
  };
  if (!connector || typeof connector !== "object") fail("is not an object.");
  if (!isSourceId(connector.id))
    fail("needs an id of letters, digits, dots, dashes or underscores.");
  if (typeof connector.label !== "string" || !connector.label.trim())
    fail("needs a label.");
  if (!externalSourceKinds.includes(connector.kind))
    fail(`kind must be one of ${externalSourceKinds.join(", ")}.`);
  if (typeof connector.read !== "function") fail("needs a read function.");
  const columns = connector.columns;
  if (
    !Array.isArray(columns) ||
    columns.length < 1 ||
    columns.length > externalSourceLimits.columns
  )
    fail(`needs 1 to ${externalSourceLimits.columns} columns.`);
  const names = new Set();
  for (const column of columns) {
    if (
      !column ||
      typeof column.name !== "string" ||
      !column.name ||
      column.name !== column.name.trim() ||
      column.name.length > 128
    )
      fail("has a column without a usable name.");
    if (!columnTypes.has(column.type))
      fail(`column "${column.name}" needs type text, number, boolean or date.`);
    if (names.has(column.name.toLowerCase()))
      fail(`column "${column.name}" is listed twice.`);
    names.add(column.name.toLowerCase());
  }
  const controlFields = connector.controlFields ?? [];
  for (const field of controlFields)
    if (
      !columns.some(
        (column) => column.name === field && column.type === "number",
      )
    )
      fail(`control field "${field}" must be a number column.`);
  const table = {
    id: connector.id,
    label: connector.label.trim(),
    kind: connector.kind,
    columns: columns.map((column) => ({
      name: column.name,
      type: column.type,
      nullable: column.nullable === true,
    })),
    controlFields: [...controlFields],
  };
  // The schema version changes exactly when the columns or control fields do.
  const schemaVersion =
    connector.schemaVersion ??
    `s-${sha256Hex(JSON.stringify([table.columns, table.controlFields])).slice(0, 16)}`;
  if (!isSourceId(schemaVersion))
    fail("schemaVersion is not a valid identifier.");
  const source = { ...table, schemaVersion };
  // Last, the editor's own schema: a source it would reject must not be registered.
  try {
    parseActiveSourceApproval(
      { source, expiresAt: "9999-12-31T00:00:00.000Z" },
      0,
    );
  } catch {
    fail(
      "is not a source description the editor accepts (check the label, column names and control fields).",
    );
  }
  return source;
}

/**
 * Creates a gateway over a fixed set of connectors.
 *
 * @param options.connectors   The sources this server offers (see src/sources/connectors.js).
 * @param options.isAllowed    `(sourceId, principal) => boolean | Promise<boolean>`: whether this
 *   caller may see the source. REQUIRED: there is no default that allows access. Pass
 *   `() => true` only for a server that has no users to tell apart.
 * @param options.approvalTtlMs  How long an approval lasts. Default 15 minutes; at most 24 hours.
 * @param options.now          Clock, for tests. Default `Date.now`.
 * @returns `{ catalog, authorize, refresh, handle }`. Each takes `{ principal, signal }`;
 *   `principal` is whatever your authentication produced and is passed to `isAllowed` unchanged.
 */
export function createSourceGateway(options = {}) {
  const { connectors, isAllowed } = options;
  if (typeof isAllowed !== "function")
    throw new TypeError(
      "createSourceGateway needs isAllowed(sourceId, principal): there is no default that allows access.",
    );
  if (
    !Array.isArray(connectors) ||
    connectors.length < 1 ||
    connectors.length > externalSourceLimits.sources
  )
    throw new TypeError(
      `createSourceGateway needs 1 to ${externalSourceLimits.sources} connectors.`,
    );
  const approvalTtlMs = options.approvalTtlMs ?? 15 * 60 * 1000;
  if (
    !Number.isInteger(approvalTtlMs) ||
    approvalTtlMs < 1000 ||
    approvalTtlMs > 24 * 60 * 60 * 1000
  )
    throw new TypeError(
      "approvalTtlMs must be between one second and 24 hours.",
    );
  const now = options.now ?? Date.now;

  const registered = new Map();
  for (const connector of connectors) {
    const source = describeSource(connector);
    if (registered.has(source.id))
      throw new TypeError(`Source id "${source.id}" is registered twice.`);
    registered.set(source.id, { connector, source });
  }

  const approvalFor = (source) => ({
    source,
    expiresAt: new Date(now() + approvalTtlMs).toISOString(),
  });
  async function allowed(sourceId, principal) {
    try {
      return (await isAllowed(sourceId, principal)) === true;
    } catch {
      return false;
    }
  }
  /** The registered source if the caller may see it. An unknown id is refused the same way. */
  async function resolve(sourceId, principal) {
    const entry = isSourceId(sourceId) ? registered.get(sourceId) : undefined;
    if (!entry || !(await allowed(sourceId, principal)))
      throw new ExternalSourceError("refused");
    return entry;
  }

  async function catalog({ principal } = {}) {
    const approvals = [];
    for (const { source } of registered.values())
      if (await allowed(source.id, principal))
        approvals.push(approvalFor(source));
    return approvals;
  }

  async function authorize(sourceId, { principal } = {}) {
    const { source } = await resolve(sourceId, principal);
    return approvalFor(source);
  }

  async function refresh(sourceId, { principal, signal } = {}) {
    const { connector, source } = await resolve(sourceId, principal);
    if (signal?.aborted) throw new ExternalSourceError("cancelled");
    let result;
    try {
      result = await connector.read({ signal });
    } catch (error) {
      throw toExternalSourceError(error);
    }
    if (!result || !Array.isArray(result.rows))
      throw new ExternalSourceError("invalid");
    const timestamp =
      result.sourceTimestamp instanceof Date
        ? result.sourceTimestamp
        : new Date(result.sourceTimestamp ?? now());
    if (!Number.isFinite(timestamp.getTime()))
      throw new ExternalSourceError("invalid");
    const payload = {
      protocol: externalSourceFormatVersion,
      source,
      sourceTimestamp: timestamp.toISOString(),
      // Identifies this exact content, so the editor can tell a changed source from an unchanged one.
      version:
        result.version ??
        `v-${sha256Hex(JSON.stringify(result.rows)).slice(0, 32)}`,
      rows: result.rows,
      complete: true,
    };
    // The editor's own check: shape, limits, column types, and a match with the approved source.
    return validateExternalSourcePayload(payload, source);
  }

  /** Runs one protocol request, as parsed from a request body. */
  async function handle(body, context = {}) {
    const request = parseSourceRequest(body);
    if (request.operation === "catalog") return catalog(context);
    if (request.operation === "authorize")
      return authorize(request.sourceId, context);
    return refresh(request.sourceId, context);
  }

  return { catalog, authorize, refresh, handle };
}

/** HTTP status for each protocol error code. The body always carries the code. */
export const sourceErrorStatus = Object.freeze({
  invalid: 400,
  refused: 403,
  expired: 403,
  stale: 409,
  schema: 422,
  limit: 413,
  timeout: 504,
  cancelled: 499,
  unavailable: 503,
});

/**
 * Wraps a gateway as a Node HTTP request handler for one POST endpoint.
 *
 *   POST <path>  body: { "operation": "catalog" }            -> 200 [approval, ...]
 *                body: { "operation": "authorize", sourceId } -> 200 approval
 *                body: { "operation": "refresh", sourceId }   -> 200 payload
 *   errors                                                    -> { "error": { "code", "message" } }
 *
 * @param options.authenticate  `(request) => principal | Promise<principal>`. REQUIRED. Return the
 *   caller's identity for `isAllowed`, or throw / return null to refuse the request (401).
 * @param options.path          Only this path is served. Default "/api/sources".
 * @param options.maxBodyBytes  Default 4096; a protocol request is tiny.
 * @param options.timeoutMs     Per-request limit. Default and maximum: the protocol's 30 seconds.
 * @returns `async (request, response) => boolean`: true if it handled the request, false if the
 *   path is not this endpoint (so it can sit beside other routes).
 */
export function createSourceGatewayHttpHandler(gateway, options = {}) {
  if (typeof options.authenticate !== "function")
    throw new TypeError(
      "createSourceGatewayHttpHandler needs authenticate(request): there is no default that allows access.",
    );
  const path = options.path ?? "/api/sources";
  const maxBodyBytes = options.maxBodyBytes ?? 4096;
  const timeoutMs = Math.min(
    options.timeoutMs ?? externalSourceLimits.timeoutMs,
    externalSourceLimits.timeoutMs,
  );

  const send = (response, status, value) => {
    const body = JSON.stringify(value);
    response.writeHead(status, {
      "Content-Type": "application/json; charset=utf-8",
      "Content-Length": Buffer.byteLength(body),
      "Cache-Control": "no-store",
      "X-Content-Type-Options": "nosniff",
    });
    response.end(body);
  };
  const sendError = (response, error) => {
    const sourceError = toExternalSourceError(error);
    send(response, sourceErrorStatus[sourceError.code] ?? 503, {
      error: { code: sourceError.code, message: sourceError.message },
    });
  };

  return async function handleSourceRequest(request, response) {
    const url = new URL(request.url ?? "/", "http://localhost");
    if (url.pathname !== path) return false;
    if (request.method !== "POST") {
      response.writeHead(405, { Allow: "POST" });
      response.end();
      return true;
    }
    if (!/^application\/json\b/i.test(request.headers["content-type"] ?? "")) {
      send(response, 415, {
        error: { code: "invalid", message: "Send application/json." },
      });
      return true;
    }
    let principal;
    try {
      principal = await options.authenticate(request);
    } catch {
      principal = null;
    }
    if (principal === null || principal === undefined || principal === false) {
      send(response, 401, {
        error: { code: "refused", message: "Sign in to use this source." },
      });
      return true;
    }
    const abort = new AbortController();
    const timer = setTimeout(
      () => abort.abort(new ExternalSourceError("timeout")),
      timeoutMs,
    );
    const onClose = () => {
      if (!response.writableEnded)
        abort.abort(new ExternalSourceError("cancelled"));
    };
    response.on("close", onClose);
    try {
      const chunks = [];
      let size = 0;
      for await (const chunk of request) {
        size += chunk.length;
        if (size > maxBodyBytes) throw new ExternalSourceError("invalid");
        chunks.push(chunk);
      }
      let body;
      try {
        body = JSON.parse(Buffer.concat(chunks).toString("utf8"));
      } catch {
        throw new ExternalSourceError("invalid");
      }
      const result = await Promise.race([
        gateway.handle(body, { principal, signal: abort.signal }),
        new Promise((_, reject) =>
          abort.signal.addEventListener(
            "abort",
            () => reject(abort.signal.reason),
            {
              once: true,
            },
          ),
        ),
      ]);
      send(response, 200, result);
    } catch (error) {
      if (!response.writableEnded && !response.destroyed)
        sendError(response, error);
    } finally {
      clearTimeout(timer);
      response.off("close", onClose);
    }
    return true;
  };
}
