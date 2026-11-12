import type { Workbook } from "../index";

/** Why a file was refused or could not be read. */
export type FileProcessingErrorCode =
  | "FILE_TYPE"
  | "FILE_INVALID"
  | "FILE_LIMIT"
  | "FILE_ACTIVE"
  | "FILE_CANCELLED"
  | "FILE_BUSY"
  | "FILE_WORKER"
  | (string & {});

export interface FileProcessingErrorInstance extends Error {
  name: "FileProcessingError";
  /** For example "FILE_TYPE" (extension not accepted) or "FILE_INVALID" (not a readable file of that type). */
  code: FileProcessingErrorCode;
}
/** The error every file function throws. Its message is written for the person who chose the file. */
export const FileProcessingError: new (
  code: FileProcessingErrorCode,
  message?: string,
) => FileProcessingErrorInstance;

export interface FileImportOptions {
  /** Cancels the import; the promise then rejects with code "FILE_CANCELLED". */
  signal?: AbortSignal;
  /** Gives up after this many milliseconds. */
  timeoutMs?: number;
  [extra: string]: unknown;
}

/** How delimited text is read. Unset fields take the value in `delimitedImportDefaults`. */
export interface DelimitedImportOptions {
  /** "auto" detects comma, semicolon, tab or pipe. */
  delimiter?: "auto" | string;
  encoding?: "auto" | string;
  quote?: string;
  escape?: string;
  newline?: "auto" | string;
  /** True when the first record holds column names. */
  hasHeader?: boolean;
  /** Decimal mark in the file: "." or ",". */
  decimal?: string;
  grouping?: string;
  /** Order of day, month and year in text dates, for example "ymd" or "dmy". */
  dateOrder?: string;
  empty?: string;
  trim?: boolean;
  /**
   * Column types by zero-based column index: "auto", "text", "number" and others. A decimal comma or
   * a grouping separator is only applied to columns typed "number"; "auto" columns keep such values as text.
   */
  columns?: Record<string, string>;
}
/** The values `importDelimitedFile` uses for any option left out. */
export const delimitedImportDefaults: Readonly<
  Required<DelimitedImportOptions>
>;

export interface WorkbookImportResult {
  workbook: Workbook;
  /** What the format could not carry over, in plain sentences. */
  warnings: string[];
  report: Record<string, unknown>;
}

/**
 * Reads an XLSX workbook (or another workbook format the importer accepts) from a `File`.
 * Parsing runs in a worker. Rejects with `FileProcessingError` if the file is not accepted.
 */
export function importWorkbookFile(
  file: File,
  importOptions?: FileImportOptions,
): Promise<WorkbookImportResult>;

/**
 * Reads delimited text (CSV, TSV) into a one-sheet workbook. The report lists the detected
 * settings, the column types and anything that was changed while reading. Text that begins with "="
 * is stored as text (`literal: true`), never as a formula.
 */
export function importDelimitedFile(
  file: File,
  parserOptions?: DelimitedImportOptions,
  importOptions?: FileImportOptions,
): Promise<{ workbook: Workbook; report: Record<string, unknown> }>;

/**
 * Parses a native JSON file (see `encodeNativeWorkbook`) in a worker, within the file size limits.
 * It does not check that the content is a workbook: pass the result to `normalizeWorkbookInput`,
 * which validates it and throws if it is not one.
 */
export function importNativeJsonFile(
  file: File,
  importOptions?: FileImportOptions,
): Promise<unknown>;

/**
 * Builds an XLSX file from the workbook. `warnings` lists what XLSX cannot keep (for example
 * review history and pivot definitions); native JSON keeps everything.
 */
export function exportWorkbookToXlsx(workbook: Workbook): Promise<{
  blob: Blob;
  warnings: string[];
  report: Record<string, unknown>;
}>;

export interface DelimitedExportOptions {
  /** Defaults to ",". */
  delimiter?: string;
  /** Defaults to "utf-8". */
  encoding?: string;
  /** Defaults to "\r\n". */
  newline?: string;
  includeHeader?: boolean;
  /** Writes a byte-order mark. Defaults to false. */
  bom?: boolean;
  /** Quotes every field. Defaults to false. */
  quoteAll?: boolean;
  /** Range to write, such as "A1:D20". Defaults to the used range. */
  range?: string;
  [extra: string]: unknown;
}

/**
 * Writes one sheet as delimited text, with formulas replaced by their calculated values.
 * Throws if the sheet does not exist or an option is not valid.
 */
export function exportWorksheetToCsv(
  workbook: Workbook,
  sheetId: string,
  options?: DelimitedExportOptions,
): { text: string; bytes: Uint8Array; report: Record<string, unknown> };

/** The workbook as native JSON text: the format that keeps everything. Read it back with `importNativeJsonFile`. */
export function encodeNativeWorkbook(workbook: Workbook): string;
