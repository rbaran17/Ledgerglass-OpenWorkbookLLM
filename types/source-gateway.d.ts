/**
 * Type declarations for `@ledgerglass/workbook/source-gateway`: the server side of the external
 * source protocol, `ledgerglass-sources-v1`. For Node servers only. See docs/SOURCES.md.
 */
import type { IncomingMessage, ServerResponse } from "node:http";

/** The protocol's format name. The editor refuses a reply that carries any other value. */
export const externalSourceFormatVersion: "ledgerglass-sources-v1";

/** What a source is, as the editor lists it. */
export type SourceKind =
  | "database"
  | "odbc"
  | "odata"
  | "api"
  | "erp"
  | "file-folder";

/** The kinds a source may have, in the order the protocol lists them. */
export const externalSourceKinds: readonly SourceKind[];

/** The nine error codes the protocol carries. Any other failure reaches the user as "unavailable". */
export type ExternalSourceErrorCode =
  | "unavailable"
  | "refused"
  | "expired"
  | "schema"
  | "limit"
  | "timeout"
  | "cancelled"
  | "invalid"
  | "stale";

/** The nine error codes, in the order the protocol lists them. */
export const externalSourceErrorCodes: readonly ExternalSourceErrorCode[];

export interface ExternalSourceErrorInstance extends Error {
  name: "ExternalSourceError";
  /** The protocol code. The message is the text the editor shows for that code. */
  code: ExternalSourceErrorCode;
}

/** The error the protocol raises. Throws nothing itself; construct it with a code from `ExternalSourceErrorCode`. */
export const ExternalSourceError: new (
  code: ExternalSourceErrorCode,
) => ExternalSourceErrorInstance;

/**
 * The published limits. The editor and the gateway both enforce them; a value over a limit is
 * refused, never truncated.
 */
export const externalSourceLimits: Readonly<{
  /** Most sources one gateway offers, and most approvals in one catalog. */
  sources: number;
  /** Most columns in one source. */
  columns: number;
  /** Most data rows in one refresh, not counting the header. */
  rows: number;
  /** Most cells in one refresh, counting the header row: (rows + 1) x columns. */
  cells: number;
  /** Most bytes in one reply, and in one source file read by a file connector. */
  responseBytes: number;
  /** Longest a single protocol operation may take, in milliseconds. */
  timeoutMs: number;
  /** Most characters of saved connection state in one workbook, counted as JSON across its sheets. */
  metadataCharacters: number;
}>;

/** The column types a source may declare. */
export type SourceColumnType = "text" | "number" | "boolean" | "date";

/** One column as a connector declares it. Leaving `nullable` out means the column may not be empty. */
export interface SourceColumn {
  /** Unique within the source, ignoring case. Up to 128 characters, with no leading or trailing space. */
  name: string;
  type: SourceColumnType;
  nullable?: boolean;
}

/** A source's description, as the editor receives it. The gateway fills `nullable` and `schemaVersion`. */
export interface SourceTable {
  /** The identifier the editor and the host use for the source. */
  id: string;
  label: string;
  kind: SourceKind;
  /** Changes exactly when the columns or control fields change. */
  schemaVersion: string;
  columns: Array<SourceColumn & { nullable: boolean }>;
  /** Numeric columns whose totals the editor compares before a refresh is applied. */
  controlFields: string[];
}

/** One source the caller may use, with the time until which that approval is valid. */
export interface SourceApproval {
  source: SourceTable;
  /** An ISO 8601 date-time with a time zone. The editor refuses an approval after this time. */
  expiresAt: string;
}

/** One cell in a refresh: text, a finite number, a boolean, or null for an empty value. */
export type SourceCellValue = string | number | boolean | null;

/** The reply to a refresh: the source's rows, with the source's description. */
export interface SourcePayload {
  protocol: "ledgerglass-sources-v1";
  source: SourceTable;
  /** An ISO 8601 date-time with a time zone. Set from the connector, or from the time of the read. */
  sourceTimestamp: string;
  /** Identifies this exact content. A changed source has a different version. */
  version: string;
  /** Values in column order. */
  rows: SourceCellValue[][];
  complete: true;
}

/** What a connector's `read` resolves to. */
export interface SourceReadResult {
  /** Values in column order, already typed to the declared columns. */
  rows: SourceCellValue[][];
  /** When the data was last changed at the source. Default: the time of the read. */
  sourceTimestamp?: Date | string | number;
  /** Identifies this content. Default: a hash of the rows. Must be a valid identifier. */
  version?: string;
}

/**
 * A source the gateway serves: a plain object with a description and a `read` function. The
 * connectors in `@ledgerglass/workbook/source-connectors` return this shape.
 */
export interface SourceConnector {
  /** Letters, digits, dots, dashes or underscores; 1 to 100 characters. Unique within the gateway. */
  id: string;
  label: string;
  kind: SourceKind;
  columns: SourceColumn[];
  /** Names of numeric columns to total as control fields. */
  controlFields?: string[];
  /** Default: a hash of the columns and control fields. */
  schemaVersion?: string;
  /** Reads the source. Rejects with an error to report a failure; an `ExternalSourceError` keeps its code. */
  read(context: { signal?: AbortSignal }): Promise<SourceReadResult> | SourceReadResult;
}

/** One protocol request, as the editor sends it. */
export type SourceRequest =
  | { operation: "catalog" }
  | { operation: "authorize" | "refresh"; sourceId: string };

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
