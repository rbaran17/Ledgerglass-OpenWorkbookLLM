/**
 * File helpers for Node: read a workbook from disk and write one back, by path. They wrap the
 * same import and export functions the browser uses.
 */
import { readFile, writeFile } from "node:fs/promises";
import { basename, extname } from "node:path";
import {
  importDelimitedFile,
  importNativeJsonFile,
  importWorkbookFile,
} from "../app/07-file-import-and-worker-dispatch.js";
import {
  encodeNativeWorkbook,
  exportWorkbookToXlsx,
  exportWorksheetToCsv,
} from "../app/08-xlsx-import-export-and-preservation.js";
import { normalizeWorkbookInput } from "../app/03-workbook-model.js";

const delimitedExtensions = new Set([".csv", ".tsv", ".txt"]);

/**
 * Reads a workbook file. The format follows the extension: .json is native JSON, .csv/.tsv/.txt
 * are delimited text, anything else goes through the workbook importer (XLSX and the other
 * formats it accepts).
 *
 * @param path  File to read.
 * @param options.delimited  Parser options for delimited text (see `delimitedImportDefaults`).
 * @param options.signal, options.timeoutMs  Passed to the importer.
 * @returns `{ workbook, warnings, report }`; `warnings` is empty and `report` undefined where the
 *   format has none.
 * @throws FileProcessingError for a file that is unsupported, unsafe, too large or invalid; an
 *   Error for native JSON that is not a workbook.
 */
export async function readWorkbookFile(path, options = {}) {
  const bytes = await readFile(path);
  const file = new File([bytes], basename(path));
  const extension = extname(path).toLowerCase();
  const importOptions = {
    ...(options.signal ? { signal: options.signal } : {}),
    ...(options.timeoutMs ? { timeoutMs: options.timeoutMs } : {}),
  };
  if (extension === ".json") {
    const workbook = normalizeWorkbookInput(
      await importNativeJsonFile(file, importOptions),
    );
    return { workbook, warnings: [], report: undefined };
  }
  if (delimitedExtensions.has(extension)) {
    const result = await importDelimitedFile(
      file,
      options.delimited ?? {},
      importOptions,
    );
    return { workbook: result.workbook, warnings: [], report: result.report };
  }
  const result = await importWorkbookFile(file, importOptions);
  return {
    workbook: result.workbook,
    warnings: result.warnings ?? [],
    report: result.report,
  };
}

/**
 * Writes a workbook. The format follows the extension: .xlsx, .json (native, keeps everything) or
 * .csv/.tsv (one sheet).
 *
 * @param options.sheetId  Sheet to write for delimited text. Defaults to the first sheet.
 * @param options.delimited  Options for delimited text (delimiter, includeHeader, bom, ...).
 *   The delimiter defaults to a tab for .tsv.
 * @returns `{ path, bytes, warnings }`; `warnings` lists what the format could not keep.
 */
export async function writeWorkbookFile(workbook, path, options = {}) {
  const extension = extname(path).toLowerCase();
  let data;
  let warnings = [];
  if (extension === ".xlsx") {
    const result = await exportWorkbookToXlsx(workbook);
    data = Buffer.from(await result.blob.arrayBuffer());
    warnings = result.warnings ?? [];
  } else if (extension === ".json") {
    data = Buffer.from(encodeNativeWorkbook(normalizeWorkbookInput(workbook)));
  } else if (extension === ".csv" || extension === ".tsv") {
    const sheetId = options.sheetId ?? workbook.sheets[0]?.id;
    const result = exportWorksheetToCsv(workbook, sheetId, {
      ...(extension === ".tsv" ? { delimiter: "\t" } : {}),
      ...options.delimited,
    });
    data = Buffer.from(result.bytes);
    warnings = result.warnings ?? [];
  } else
    throw new TypeError(
      `Cannot write "${extension || path}": use .xlsx, .json, .csv or .tsv.`,
    );
  await writeFile(path, data);
  return { path, bytes: data.length, warnings };
}
