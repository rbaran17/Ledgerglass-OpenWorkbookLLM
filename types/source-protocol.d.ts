/**
 * Types and constants of the external source protocol ("ledgerglass-sources-v1"). They are shared
 * by the browser entry and the server entries and use no Node types.
 */
/** The protocol's format name. The editor refuses a reply that carries any other value. */
export const externalSourceFormatVersion: "ledgerglass-sources-v1";

/** What a source is, as the editor lists it. */
export type SourceKind =
  "database" | "odbc" | "odata" | "api" | "erp" | "file-folder";

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
  read(context: {
    signal?: AbortSignal;
  }): Promise<SourceReadResult> | SourceReadResult;
}

/** One protocol request, as the editor sends it. */
export type SourceRequest =
  | { operation: "catalog" }
  | { operation: "authorize" | "refresh"; sourceId: string };
