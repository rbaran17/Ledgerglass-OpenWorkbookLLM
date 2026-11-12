/**
 * Source connectors: each one reads rows from one place and hands them to the gateway in the
 * declared column order and types. A connector is a plain object
 * `{ id, label, kind, columns, controlFields?, read({ signal }) }`; `read` resolves to
 * `{ rows, sourceTimestamp?, version? }`. Write your own in that shape for anything not here.
 *
 * Everything a connector reaches (file, URL, query) is fixed by the host when it is created.
 * Nothing in a browser request can change it.
 *
 * New code: the original product's connectors were server-side and were not recovered.
 */
import { readFile, realpath, stat } from "node:fs/promises";
import { isAbsolute, relative, resolve, sep } from "node:path";
import Papa from "papaparse";
import {
  ExternalSourceError,
  externalSourceLimits,
  shapeSourceRows,
  toExternalSourceError,
} from "./protocol.js";

const maxSourceFileBytes = externalSourceLimits.responseBytes;

function requireText(value, name, connectorName) {
  if (typeof value !== "string" || !value)
    throw new TypeError(`${connectorName} needs ${name}.`);
  return value;
}

/**
 * One delimited text file (CSV, TSV) inside a folder the host names.
 *
 * @param options.directory  The folder. The file must resolve inside it, also through links.
 * @param options.file       File name relative to the folder.
 * @param options.columns    Declared columns, in file order. The header row must match their names.
 * @param options.delimiter  Default: detected. Set "," or "\t" or ";" to be exact.
 * @param options.hasHeader  Default true. With false, columns are taken by position.
 */
export function createDelimitedFileConnector(options = {}) {
  const name = "createDelimitedFileConnector";
  const directory = resolve(
    requireText(options.directory, "a directory", name),
  );
  const file = requireText(options.file, "a file", name);
  if (isAbsolute(file) || file.split(/[\\/]/).includes(".."))
    throw new TypeError(`${name}: file must be a name inside the directory.`);
  const hasHeader = options.hasHeader !== false;
  return {
    id: options.id,
    label: options.label,
    kind: "file-folder",
    columns: options.columns,
    controlFields: options.controlFields,
    async read() {
      let path;
      let information;
      try {
        // Resolve links before checking the folder, so a link cannot point outside it.
        const root = await realpath(directory);
        path = await realpath(resolve(root, file));
        const inside = relative(root, path);
        if (
          !inside ||
          inside.startsWith(`..${sep}`) ||
          inside === ".." ||
          isAbsolute(inside)
        )
          throw new ExternalSourceError("refused");
        information = await stat(path);
      } catch (error) {
        throw toExternalSourceError(error);
      }
      if (!information.isFile()) throw new ExternalSourceError("unavailable");
      if (information.size > maxSourceFileBytes)
        throw new ExternalSourceError("limit");
      const text = (await readFile(path, "utf8")).replace(/^﻿/, "");
      const parsed = Papa.parse(text, {
        delimiter: options.delimiter ?? "",
        skipEmptyLines: "greedy",
      });
      if (parsed.errors.some((problem) => problem.type !== "Delimiter"))
        throw new ExternalSourceError("invalid");
      const records = parsed.data;
      if (hasHeader) {
        const header = (records.shift() ?? []).map((cell) =>
          String(cell).trim(),
        );
        const expected = options.columns.map((column) => column.name);
        if (
          header.length !== expected.length ||
          header.some((cell, index) => cell !== expected[index])
        )
          throw new ExternalSourceError("schema");
      }
      if (records.some((record) => record.length !== options.columns.length))
        throw new ExternalSourceError("schema");
      return {
        rows: shapeSourceRows(records, options.columns, { parseText: true }),
        sourceTimestamp: information.mtime,
      };
    },
  };
}

function readPath(value, path) {
  if (!path) return value;
  let current = value;
  for (const key of path.split("."))
    current =
      current === null || current === undefined ? undefined : current[key];
  return current;
}

async function fetchJson(fetchImplementation, url, headers, signal) {
  let response;
  try {
    response = await fetchImplementation(url, {
      method: "GET",
      headers: { Accept: "application/json", ...headers },
      redirect: "error",
      signal,
    });
  } catch (error) {
    throw toExternalSourceError(error);
  }
  if (response.status === 401 || response.status === 403)
    throw new ExternalSourceError("refused");
  if (!response.ok) throw new ExternalSourceError("unavailable");
  const declared = Number(response.headers?.get?.("content-length"));
  if (
    Number.isFinite(declared) &&
    declared > externalSourceLimits.responseBytes
  )
    throw new ExternalSourceError("limit");
  const text = await response.text();
  if (Buffer.byteLength(text) > externalSourceLimits.responseBytes)
    throw new ExternalSourceError("limit");
  try {
    return JSON.parse(text);
  } catch {
    throw new ExternalSourceError("invalid");
  }
}

function requireHttpUrl(value, connectorName) {
  let url;
  try {
    url = new URL(value);
  } catch {
    throw new TypeError(`${connectorName} needs an absolute URL.`);
  }
  if (url.protocol !== "https:" && url.protocol !== "http:")
    throw new TypeError(`${connectorName}: the URL must be http or https.`);
  return url;
}

/**
 * A JSON endpoint that returns a list of records.
 *
 * @param options.url       The endpoint. Fixed here; a request cannot change it.
 * @param options.headers   For example an Authorization header. Kept on the server.
 * @param options.rowsPath  Dot path to the list in the reply, for example "data.items".
 *   Default: the reply itself is the list.
 * @param options.fields    Record key for each column, in column order. Default: the column names.
 * @param options.fetch     Default: the global fetch. Redirects are refused.
 */
export function createJsonApiConnector(options = {}) {
  const url = requireHttpUrl(options.url, "createJsonApiConnector");
  const fetchImplementation = options.fetch ?? globalThis.fetch;
  return {
    id: options.id,
    label: options.label,
    kind: options.kind ?? "api",
    columns: options.columns,
    controlFields: options.controlFields,
    async read({ signal } = {}) {
      const reply = await fetchJson(
        fetchImplementation,
        url,
        options.headers,
        signal,
      );
      const records = readPath(reply, options.rowsPath);
      if (!Array.isArray(records)) throw new ExternalSourceError("invalid");
      const fields =
        options.fields ?? options.columns.map((column) => column.name);
      return {
        rows: shapeSourceRows(
          records.map((record) =>
            fields.map((field) => readPath(record, field)),
          ),
          options.columns,
        ),
      };
    },
  };
}

/**
 * An OData entity set. Asks only for the declared columns (`$select`) and follows
 * `@odata.nextLink` pages on the same origin until the set is complete or over the row limit.
 *
 * @param options.url      The entity set URL, for example "https://host/odata/Invoices".
 * @param options.filter   An optional fixed `$filter` expression.
 * @param options.fields   Property name for each column. Default: the column names.
 */
export function createODataConnector(options = {}) {
  const url = requireHttpUrl(options.url, "createODataConnector");
  const fetchImplementation = options.fetch ?? globalThis.fetch;
  const fields =
    options.fields ?? options.columns?.map((column) => column.name);
  return {
    id: options.id,
    label: options.label,
    kind: "odata",
    columns: options.columns,
    controlFields: options.controlFields,
    async read({ signal } = {}) {
      // Written by hand so the option names keep their "$" and spaces travel as %20, which is
      // what OData services expect; URLSearchParams would send "%24select" and "+".
      const first = new URL(url);
      const query = [
        first.search.replace(/^\?/, ""),
        `$select=${fields.map(encodeURIComponent).join(",")}`,
        options.filter ? `$filter=${encodeURIComponent(options.filter)}` : "",
      ].filter(Boolean);
      first.search = query.join("&");
      const records = [];
      let next = first;
      // A page link is followed only on the origin the host configured.
      for (let page = 0; next; page++) {
        if (page >= 200) throw new ExternalSourceError("limit");
        const reply = await fetchJson(
          fetchImplementation,
          next,
          options.headers,
          signal,
        );
        if (!Array.isArray(reply?.value))
          throw new ExternalSourceError("invalid");
        records.push(...reply.value);
        if (records.length > externalSourceLimits.rows)
          throw new ExternalSourceError("limit");
        const link = reply["@odata.nextLink"];
        if (!link) break;
        let following;
        try {
          following = new URL(link, next);
        } catch {
          throw new ExternalSourceError("invalid");
        }
        if (following.origin !== url.origin)
          throw new ExternalSourceError("refused");
        next = following;
      }
      return {
        rows: shapeSourceRows(
          records.map((record) => fields.map((field) => record?.[field])),
          options.columns,
        ),
      };
    },
  };
}

/**
 * A fixed query, run by a function the host supplies. Use it for any database driver, an ODBC
 * connection or an ERP client: the driver stays in your application, and this library needs none.
 *
 * @param options.kind  "database", "odbc" or "erp". Default "database".
 * @param options.run   `({ signal }) => rows | Promise<rows>`: arrays in column order, or objects
 *   keyed by column name. Build the query here; nothing from a request reaches it.
 */
export function createQueryConnector(options = {}) {
  if (typeof options.run !== "function")
    throw new TypeError("createQueryConnector needs a run function.");
  return {
    id: options.id,
    label: options.label,
    kind: options.kind ?? "database",
    columns: options.columns,
    controlFields: options.controlFields,
    async read({ signal } = {}) {
      let records;
      try {
        records = await options.run({ signal });
      } catch (error) {
        throw toExternalSourceError(error);
      }
      return { rows: shapeSourceRows(records, options.columns) };
    },
  };
}

/**
 * A fixed SELECT on a SQLite database file, opened read-only, through Node's built-in
 * `node:sqlite` (Node 22.5 or later; Node marks that module experimental).
 *
 * @param options.database  Path to the database file.
 * @param options.query     One SELECT statement whose result columns are the declared columns,
 *   in order. Fixed here; a request cannot change it. Anything else is refused when the
 *   connector is created, and the read-only connection refuses writes in any case.
 * @param options.parameters  Values for `?` placeholders in the query. Also fixed here.
 */
export function createSqliteConnector(options = {}) {
  const name = "createSqliteConnector";
  const database = resolve(
    requireText(options.database, "a database path", name),
  );
  const query = requireText(options.query, "a query", name).trim();
  // A first check on the text, with string literals, quoted names and comments removed. The
  // read-only connection below is what enforces it: SQLite itself refuses any write.
  const bare = query
    .replace(/'(?:[^']|'')*'/g, "''")
    .replace(/"(?:[^"]|"")*"/g, '""')
    .replace(/--[^\n]*/g, " ")
    .replace(/\/\*[\s\S]*?\*\//g, " ");
  if (
    !/^(select|with)\b/i.test(bare) ||
    /;\s*\S/.test(bare) ||
    /\b(insert|update|delete|replace|drop|alter|create|attach|detach|pragma|vacuum|reindex)\b/i.test(
      bare,
    )
  )
    throw new TypeError(`${name}: query must be a single SELECT statement.`);
  return {
    id: options.id,
    label: options.label,
    kind: "database",
    columns: options.columns,
    controlFields: options.controlFields,
    async read() {
      let handle;
      try {
        const { DatabaseSync } = await import("node:sqlite");
        const information = await stat(database);
        handle = new DatabaseSync(database, { readOnly: true });
        const statement = handle.prepare(query);
        const records = [];
        for (const record of statement.iterate(...(options.parameters ?? []))) {
          records.push(Object.values(record));
          if (records.length > externalSourceLimits.rows)
            throw new ExternalSourceError("limit");
        }
        return {
          rows: shapeSourceRows(records, options.columns),
          sourceTimestamp: information.mtime,
        };
      } catch (error) {
        throw toExternalSourceError(error);
      } finally {
        handle?.close();
      }
    },
  };
}
