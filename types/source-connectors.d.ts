/**
 * Type declarations for `@ledgerglass/workbook/source-connectors`: connectors that read rows for
 * the source gateway. For Node servers only. See docs/SOURCES.md.
 */
import type {
  SourceColumn,
  SourceConnector,
  SourceKind,
} from "./source-gateway";

/** A record a query returns: values in column order, or an object keyed by column name. */
export type SourceRecord = readonly unknown[] | Record<string, unknown>;

/** Options every connector takes. */
export interface SourceConnectorBaseOptions {
  /** The identifier the editor uses for the source. Letters, digits, dots, dashes or underscores. */
  id: string;
  label: string;
  /** The declared columns, in the order the source returns them. */
  columns: SourceColumn[];
  /** Names of numeric columns to total as control fields. */
  controlFields?: string[];
}

export interface DelimitedFileConnectorOptions extends SourceConnectorBaseOptions {
  /** The folder. The file must resolve inside it, also through links. */
  directory: string;
  /** The file's name inside the folder. Not absolute, and it may not contain "..". */
  file: string;
  /** Set "," , "\t" or ";" to fix the delimiter. Default: detected. */
  delimiter?: string;
  /** Whether the first line is a header that must match the column names. Default true. */
  hasHeader?: boolean;
}

export interface JsonApiConnectorOptions extends SourceConnectorBaseOptions {
  /** The endpoint, as an absolute http or https URL. Fixed here; a request cannot change it. */
  url: string;
  /** Sent with every request, for example an Authorization header. Kept on the server. */
  headers?: Record<string, string>;
  /** Dot path to the list of records in the reply, for example "data.items". Default: the reply is the list. */
  rowsPath?: string;
  /** The record key for each column, in column order. Default: the column names. Dot paths are followed. */
  fields?: string[];
  /** The source kind. Default "api". */
  kind?: SourceKind;
  /** The fetch to use. Default: the global fetch. Redirects are refused. */
  fetch?: typeof fetch;
}

export interface ODataConnectorOptions extends SourceConnectorBaseOptions {
  /** The entity set URL, for example "https://host/odata/Invoices". Fixed here. */
  url: string;
  /** A fixed `$filter` expression, sent with every request. */
  filter?: string;
  /** The property name for each column, in column order. Default: the column names. */
  fields?: string[];
  /** Sent with every request, for example an Authorization header. Kept on the server. */
  headers?: Record<string, string>;
  /** The fetch to use. Default: the global fetch. Redirects are refused. */
  fetch?: typeof fetch;
}

export interface QueryConnectorOptions extends SourceConnectorBaseOptions {
  /** The source kind. Default "database". */
  kind?: "database" | "odbc" | "erp";
  /**
   * Runs the fixed query and returns its records. An `ExternalSourceError` keeps its code, an abort
   * or timeout maps to its own code, and any other failure is reported as "unavailable".
   */
  run(context: {
    signal?: AbortSignal;
  }): SourceRecord[] | Promise<SourceRecord[]>;
}

export interface SqliteConnectorOptions extends SourceConnectorBaseOptions {
  /** Path to the SQLite database file. It is opened read-only. */
  database: string;
  /** One SELECT (or WITH ... SELECT) statement whose result columns are the declared columns, in order. */
  query: string;
  /** Values for the `?` placeholders in the query. Fixed here. */
  parameters?: ReadonlyArray<string | number | bigint | null | Uint8Array>;
}

/**
 * One delimited text file (CSV, TSV) inside a folder the host names. Throws `TypeError` for a
 * missing directory or file option, or a file that is absolute or contains "..".
 */
export function createDelimitedFileConnector(
  options: DelimitedFileConnectorOptions,
): SourceConnector;

/**
 * A JSON endpoint that returns a list of records. Throws `TypeError` for a URL that is not an
 * absolute http or https URL.
 */
export function createJsonApiConnector(
  options: JsonApiConnectorOptions,
): SourceConnector;

/**
 * An OData entity set. Asks only for the declared fields and follows `@odata.nextLink` pages on
 * the configured origin. Throws `TypeError` for a URL that is not an absolute http or https URL.
 */
export function createODataConnector(
  options: ODataConnectorOptions,
): SourceConnector;

/**
 * A fixed query run by a function the host supplies, for any database driver, ODBC connection or
 * ERP client. Throws `TypeError` without a `run` function.
 */
export function createQueryConnector(
  options: QueryConnectorOptions,
): SourceConnector;

/**
 * A fixed SELECT on a SQLite file, opened read-only through `node:sqlite`. Throws `TypeError` for
 * a missing path or a query that is not a single SELECT or WITH statement.
 */
export function createSqliteConnector(
  options: SqliteConnectorOptions,
): SourceConnector;
