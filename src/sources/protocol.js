/**
 * The external source protocol, "ledgerglass-sources-v1", as the editor's client code defines it.
 *
 * The editor asks a host for data in three operations:
 *   catalog   -> the sources this user may see, each as an approval `{ source, expiresAt }`
 *   authorize -> a fresh approval for one source
 *   refresh   -> that source's rows: `{ protocol, source, sourceTimestamp, version, rows, complete }`
 * and rejects anything that is not exactly that shape, is over the published limits, or does not
 * match the approved source description.
 *
 * The schemas, limits, error codes and validation here are the recovered editor's own, so a
 * server that passes these checks produces replies the editor accepts. The server side of the
 * original product was never in the browser bundle; src/sources/ is a new implementation of it.
 */
export {
  ExternalSourceError,
  externalSourceFormatVersion,
  externalSourceLimits,
  parseActiveSourceApproval,
  validateExternalSourcePayload,
} from "../app/03-workbook-model.js";
import {
  ExternalSourceError,
  externalSourceLimits,
} from "../app/03-workbook-model.js";

/** What a source can be, as the editor's schema lists them. */
export const externalSourceKinds = Object.freeze([
  "database",
  "odbc",
  "odata",
  "api",
  "erp",
  "file-folder",
]);

/** Error codes the editor understands. Any other failure reaches the user as "unavailable". */
export const externalSourceErrorCodes = Object.freeze([
  "unavailable",
  "refused",
  "expired",
  "schema",
  "limit",
  "timeout",
  "cancelled",
  "invalid",
  "stale",
]);

const sourceIdPattern = /^[A-Za-z0-9][A-Za-z0-9._-]{0,99}$/;

/** True for an identifier the protocol accepts: 1 to 100 letters, digits, dots, dashes, underscores. */
export function isSourceId(value) {
  return typeof value === "string" && sourceIdPattern.test(value);
}

/**
 * Checks a request body: `{ operation: "catalog" }` or
 * `{ operation: "authorize" | "refresh", sourceId }`, with no other keys.
 * @throws ExternalSourceError("invalid")
 */
export function parseSourceRequest(body) {
  if (!body || typeof body !== "object" || Array.isArray(body))
    throw new ExternalSourceError("invalid");
  const keys = Object.keys(body);
  if (body.operation === "catalog" && keys.length === 1)
    return { operation: "catalog" };
  if (
    (body.operation === "authorize" || body.operation === "refresh") &&
    keys.length === 2 &&
    isSourceId(body.sourceId)
  )
    return { operation: body.operation, sourceId: body.sourceId };
  throw new ExternalSourceError("invalid");
}

/** Returns the error as an ExternalSourceError; anything unrecognised becomes "unavailable". */
export function toExternalSourceError(error) {
  if (error instanceof ExternalSourceError) return error;
  if (
    error instanceof Error &&
    error.name === "ExternalSourceError" &&
    externalSourceErrorCodes.includes(error.code)
  )
    return new ExternalSourceError(error.code);
  if (error?.name === "AbortError") return new ExternalSourceError("cancelled");
  if (error?.name === "TimeoutError") return new ExternalSourceError("timeout");
  return new ExternalSourceError("unavailable");
}

/**
 * Turns records into protocol rows for the declared columns, strictly: a value that is not of
 * the column's type is an error, never a guess.
 *
 * @param records  Arrays in column order, or objects keyed by column name.
 * @param columns  `[{ name, type: "text" | "number" | "boolean" | "date", nullable }]`
 * @param options.parseText  When true (text sources such as CSV), text is read as the column's
 *   type: "1200.5" as a number, "true"/"false" as a boolean, "" as null.
 * @throws ExternalSourceError("limit") over the row limit; ("schema") for a value of the wrong type.
 */
export function shapeSourceRows(records, columns, options = {}) {
  if (!Array.isArray(records)) throw new ExternalSourceError("invalid");
  if (
    records.length > externalSourceLimits.rows ||
    (records.length + 1) * columns.length > externalSourceLimits.cells
  )
    throw new ExternalSourceError("limit");
  return records.map((record) =>
    columns.map((column, index) => {
      const raw = Array.isArray(record) ? record[index] : record?.[column.name];
      const value = shapeValue(raw, column, options.parseText === true);
      if (value === null && !column.nullable)
        throw new ExternalSourceError("schema");
      return value;
    }),
  );
}

const isoDatePattern = /^\d{4}-\d{2}-\d{2}$/;
function isRealDate(text) {
  if (!isoDatePattern.test(text)) return false;
  const time = Date.parse(text);
  return (
    Number.isFinite(time) && new Date(time).toISOString().slice(0, 10) === text
  );
}

function shapeValue(raw, column, parseText) {
  if (raw === null || raw === undefined) return null;
  // In a text source an empty field is "no value". Blank space is kept as text in a text column
  // and is "no value" in a number, boolean or date column.
  if (
    parseText &&
    typeof raw === "string" &&
    (raw === "" || (column.type !== "text" && raw.trim() === ""))
  )
    return null;
  switch (column.type) {
    case "text":
      if (typeof raw === "string") {
        if (raw.length > 32767) throw new ExternalSourceError("limit");
        return raw;
      }
      break;
    case "number":
      if (typeof raw === "number" && Number.isFinite(raw)) return raw;
      if (typeof raw === "bigint" && Number.isSafeInteger(Number(raw)))
        return Number(raw);
      if (
        parseText &&
        typeof raw === "string" &&
        /^[+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?$/.test(raw.trim()) &&
        Number.isFinite(Number(raw.trim()))
      )
        return Number(raw.trim());
      break;
    case "boolean":
      if (typeof raw === "boolean") return raw;
      if (parseText && typeof raw === "string") {
        const text = raw.trim().toLowerCase();
        if (text === "true") return true;
        if (text === "false") return false;
      }
      break;
    case "date": {
      if (raw instanceof Date && Number.isFinite(raw.getTime()))
        return raw.toISOString().slice(0, 10);
      if (typeof raw === "string") {
        const text = raw.trim().slice(0, 10);
        // Accept a date, or a date-time whose date part is a real date.
        if (
          isRealDate(text) &&
          (raw.trim().length === 10 ||
            /^\d{4}-\d{2}-\d{2}[T ]/.test(raw.trim()))
        )
          return text;
      }
      break;
    }
  }
  throw new ExternalSourceError("schema");
}
