/**
 * Type declarations for `@ledgerglass/workbook/source-gateway`: the server side of the external
 * source protocol, `ledgerglass-sources-v1`. For Node servers only. See docs/SOURCES.md.
 */ import type { IncomingMessage, ServerResponse } from "node:http";
import type {
  SourceKind,
  ExternalSourceErrorCode,
  ExternalSourceErrorInstance,
  SourceColumnType,
  SourceColumn,
  SourceTable,
  SourceApproval,
  SourceCellValue,
  SourcePayload,
  SourceReadResult,
  SourceConnector,
  SourceRequest,
} from "./source-protocol";

export * from "./source-protocol";

/** The context every gateway call takes: who is calling, and the signal that cancels the call. */
export interface SourceCallContext<Principal = unknown> {
  /** What `authenticate` returned for this request. Passed to `isAllowed` unchanged. */
  principal: Principal;
  signal?: AbortSignal;
}

/** Options for `createSourceGateway`. */
export interface SourceGatewayOptions<Principal = unknown> {
  /** The sources this server offers: 1 to 20, with distinct ids. Fixed when the gateway is created. */
  connectors: SourceConnector[];
  /**
   * Whether this caller may see this source. Required: there is no default that allows access.
   * A thrown error or a rejected promise counts as "no".
   */
  isAllowed(sourceId: string, principal: Principal): boolean | Promise<boolean>;
  /** How long an approval lasts, in milliseconds: at least 1000, at most 86400000. Default 15 minutes. */
  approvalTtlMs?: number;
  /** The clock used for approval times. Default `Date.now`. */
  now?: () => number;
}

/** Options for `createSourceGatewayHttpHandler`. */
export interface SourceGatewayHttpOptions<Principal = unknown> {
  /**
   * Identifies the caller from the request. Return null to refuse with 401. Required: there is no
   * default that allows access.
   */
  authenticate(
    request: IncomingMessage,
  ): Principal | null | Promise<Principal | null>;
  /** The only path served. Default "/api/sources". */
  path?: string;
  /** Largest request body accepted, in bytes. Default 4096. */
  maxBodyBytes?: number;
  /** Per-request limit in milliseconds. Capped at the protocol's `externalSourceLimits.timeoutMs`. */
  timeoutMs?: number;
}

/** A gateway over a fixed set of connectors. */
export interface SourceGateway<Principal = unknown> {
  /** The sources this caller may see, each with a fresh approval. A source the caller may not see is left out, and no connector is read. */
  catalog(context: SourceCallContext<Principal>): Promise<SourceApproval[]>;
  /** A fresh approval for one source. Rejects with `ExternalSourceError` "refused" for an unknown or forbidden source. */
  authorize(
    sourceId: string,
    context: SourceCallContext<Principal>,
  ): Promise<SourceApproval>;
  /** Reads one source and returns the checked payload. Rejects with `ExternalSourceError` on any failure. */
  refresh(
    sourceId: string,
    context: SourceCallContext<Principal>,
  ): Promise<SourcePayload>;
  /** Runs one request as parsed from a body. Rejects with `ExternalSourceError("invalid")` for a malformed body. */
  handle(
    body: unknown,
    context: SourceCallContext<Principal>,
  ): Promise<SourceApproval[] | SourceApproval | SourcePayload>;
}

/** Checks the connectors and returns a gateway. Throws `TypeError` for a missing `isAllowed`, bad connectors or a bad TTL. */
export function createSourceGateway<Principal = unknown>(
  options: SourceGatewayOptions<Principal>,
): SourceGateway<Principal>;

/**
 * Wraps a gateway as a Node request handler for one POST path. Returns true if it answered the
 * request, false if the path is not its own. Throws `TypeError` if `authenticate` is missing.
 */
export function createSourceGatewayHttpHandler<Principal = unknown>(
  gateway: SourceGateway<Principal>,
  options: SourceGatewayHttpOptions<Principal>,
): (request: IncomingMessage, response: ServerResponse) => Promise<boolean>;

/** The HTTP status the gateway sends for each error code. */
export const sourceErrorStatus: Readonly<
  Record<ExternalSourceErrorCode, number>
>;

/** True for an identifier the protocol accepts: 1 to 100 letters, digits, dots, dashes or underscores. */
export function isSourceId(value: unknown): value is string;

/** Checks a request body and returns it. Throws `ExternalSourceError("invalid")` for any other shape. */
export function parseSourceRequest(body: unknown): SourceRequest;

/** Returns the error as an `ExternalSourceError`. Anything unrecognised becomes "unavailable". */
export function toExternalSourceError(
  error: unknown,
): ExternalSourceErrorInstance;

/**
 * Turns records into protocol rows for the declared columns, strictly: a value of the wrong type
 * is an error, never a guess. Throws `ExternalSourceError("limit")` over the row or cell limit and
 * `ExternalSourceError("schema")` for a value of the wrong type.
 */
export function shapeSourceRows(
  records: ReadonlyArray<readonly unknown[] | Record<string, unknown>>,
  columns: readonly SourceColumn[],
  options?: {
    /** When true, text is read as the column's type: "1200.5" as a number, "true" as a boolean, "" as null. */
    parseText?: boolean;
  },
): SourceCellValue[][];

/**
 * Checks a refresh payload against the editor's rules, and against `source` when given. Returns
 * the checked payload. Throws `ExternalSourceError` with "invalid", "schema" or "limit".
 */
export function validateExternalSourcePayload(
  snapshot: unknown,
  source?: SourceTable,
): SourcePayload;
