# Files

Open and save workbooks: XLSX, delimited text (CSV, TSV) and native JSON. These are the functions the editor's File menu uses. They work in a browser and in Node 22 or later. Parsing runs off the main thread, in a Web Worker in a browser and in a worker thread in Node.

Native JSON is the only format that keeps everything in a workbook. XLSX and delimited text are exchange formats: the export functions report what they could not keep.

In Node, `readWorkbookFile` and `writeWorkbookFile` from `@ledgerglass/workbook/node` wrap these functions and take a path; see [the Node entry](node.md).

## Errors

Every function here rejects or throws a `FileProcessingError`. Its `message` is written for the person who chose the file. Its `code` tells the program what happened:

| Code             | Meaning                                                              |
| ---------------- | -------------------------------------------------------------------- |
| `FILE_TYPE`      | The file name or extension is not one the importer accepts.          |
| `FILE_INVALID`   | The content is not a readable file of that type, or an option is bad |
| `FILE_LIMIT`     | The file, or what it expands to, is over a size limit.               |
| `FILE_ACTIVE`    | The file carries macros or other active content.                     |
| `FILE_CANCELLED` | The `signal` was aborted.                                            |
| `FILE_BUSY`      | Too many imports are already running.                                |
| `FILE_WORKER`    | The parser worker could not be started or stopped unexpectedly.      |

### FileProcessingError

```ts
const FileProcessingError: new (
  code: string,
  message?: string,
) => Error & { name: "FileProcessingError"; code: string };
```

Use `error instanceof FileProcessingError` or check `error.name` and `error.code`.

## Opening

### importWorkbookFile

```ts
function importWorkbookFile(
  file: File,
  importOptions?: { signal?: AbortSignal; timeoutMs?: number },
): Promise<{ workbook: Workbook; warnings: string[]; report: object }>;
```

Reads an XLSX workbook. The result holds the workbook, a list of plain-language warnings about content the importer could not carry over, and a detailed report. Values, formulas, styles and page layout are read; formulas are stored as typed and calculated by Ledgerglass. It rejects with `FILE_TYPE` for an extension it does not accept and `FILE_INVALID` for content that is not an XLSX archive.

```js
import { importWorkbookFile } from "@ledgerglass/workbook";

const { workbook, warnings } = await importWorkbookFile(fileInput.files[0]);
if (warnings.length) console.warn(warnings.join("\n"));
```

### importDelimitedFile

```ts
function importDelimitedFile(
  file: File,
  parserOptions?: DelimitedImportOptions,
  importOptions?: { signal?: AbortSignal; timeoutMs?: number },
): Promise<{ workbook: Workbook; report: object }>;
```

Reads delimited text into a workbook with one sheet. With no options it detects the delimiter and the encoding and treats the first record as a header. The report gives the settings used, the number of records and columns, and the type found for each column.

Two behaviours to know:

- Text that begins with `=` is stored as text and marked `literal: true`. A CSV file cannot introduce a formula.
- With the default column type, `"auto"`, only plain numbers with a decimal point become numbers. To read a decimal comma or grouped thousands, set the column's type to `"number"`: `{ decimal: ",", columns: { 1: "number" } }`. Column keys are zero-based.

An unsupported setting is refused with a `FileProcessingError`.

```js
const { workbook, report } = await importDelimitedFile(file, {
  delimiter: ";",
  decimal: ",",
  columns: { 1: "number" },
});
```

### delimitedImportDefaults

```ts
const delimitedImportDefaults: {
  delimiter: "auto";
  encoding: "auto";
  quote: '"';
  escape: "double";
  newline: "auto";
  hasHeader: true;
  decimal: ".";
  grouping: "none";
  dateOrder: "ymd";
  empty: "blank";
  trim: false;
  columns: {};
};
```

The values `importDelimitedFile` uses for any option you leave out. Accepted values: `decimal` is `"."` or `","`; `grouping` is `"none"`, `"."`, `","` or a space; `dateOrder` is `"ymd"`, `"dmy"` or `"mdy"`; `escape` is `"double"` or `"backslash"`; `empty` is `"blank"` or `"text"`.

### importNativeJsonFile

```ts
function importNativeJsonFile(
  file: File,
  importOptions?: { signal?: AbortSignal; timeoutMs?: number },
): Promise<unknown>;
```

Parses a native JSON file in the worker, within the file size limits, and returns the parsed value. It does not check that the value is a workbook. Pass the result to `normalizeWorkbookInput`, which validates it and throws if it is not one. Text that is not JSON is refused with a `FileProcessingError`.

```js
import {
  importNativeJsonFile,
  normalizeWorkbookInput,
} from "@ledgerglass/workbook";

const workbook = normalizeWorkbookInput(await importNativeJsonFile(file));
```

## Saving

### exportWorkbookToXlsx

```ts
function exportWorkbookToXlsx(
  workbook: Workbook,
): Promise<{ blob: Blob; warnings: string[]; report: object }>;
```

Builds an XLSX file. `blob` has the XLSX media type. `warnings` always includes general notes on what XLSX does not keep (review history, outline groups, pivot and chart definitions, the date policy) and adds more when the workbook uses those features. Formulas are written as formulas; dynamic-array formulas are written as ordinary formulas at their anchor cell.

```js
const { blob, warnings } = await exportWorkbookToXlsx(workbook);
const link = document.createElement("a");
link.href = URL.createObjectURL(blob);
link.download = `${workbook.name}.xlsx`;
link.click();
```

### exportWorksheetToCsv

```ts
function exportWorksheetToCsv(
  workbook: Workbook,
  sheetId: string,
  options?: {
    delimiter?: string; // default ","
    newline?: string; // default "\r\n"
    encoding?: string; // default "utf-8"
    includeHeader?: boolean;
    bom?: boolean; // default false
    quoteAll?: boolean; // default false
    range?: string; // default: the used range
  },
): { text: string; bytes: Uint8Array; report: object };
```

Writes one sheet as delimited text. Formulas are replaced by their calculated values. `bytes` is `text` in the chosen encoding. It throws `Worksheet not found.` for an unknown sheet and refuses options that are not valid.

```js
const { text } = exportWorksheetToCsv(workbook, workbook.sheets[0].id, {
  delimiter: ";",
  range: "A1:D20",
});
```

### encodeNativeWorkbook

```ts
function encodeNativeWorkbook(workbook: Workbook): string;
```

Returns the workbook as native JSON text, indented. This is the format that keeps everything: read it back with `importNativeJsonFile` and `normalizeWorkbookInput`.
