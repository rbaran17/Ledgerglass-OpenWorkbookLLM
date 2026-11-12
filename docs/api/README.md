# API reference

The model functions the library exports, by area. Every function listed on these pages was run before it was documented, has TypeScript declarations in `types/sdk/` and has tests in `tests/unit/sdk-*.test.js`.

| Area                          | What it covers                                                                              | Exports |
| ----------------------------- | ------------------------------------------------------------------------------------------- | ------: |
| [Files](files.md)             | Open and save XLSX, delimited text and native JSON                                          |       8 |
| [Worksheets](worksheets.md)   | Add, rename, move, hide and protect sheets; panes, merges, hidden rows; workbook properties |      25 |
| [Cells](cells.md)             | Addresses, ranges and selections; borders, number formats, hyperlinks, AutoSum              |      25 |
| [Tables](tables.md)           | Worksheet tables, filters and saved filter views, pivot tables                              |      25 |
| [Rules](rules.md)             | Data validation, conditional formatting, defined names, review threads, formula watches     |      22 |
| [Calculation](calculation.md) | Calculation settings and evidence, the calculator object, function-name catalogues, what-if |      22 |
| [Node entry](node.md)         | Files by path and all calculated values at once, for scripts and servers                    |       3 |

[All exports](exports.md) lists every public name on one page, with the entry points that export it.

The editor component, the assistant adapters and the core functions (`createBlankWorkbook`, `normalizeWorkbookInput`, `getWorkbookCalculation`, `applyCellPatchTransaction`) are described in [Using Ledgerglass as a library](../LIBRARY.md). The agent-rendered workbook is in [Showing a workbook from an agent](../AGENT-UI.md), and the MCP server in [MCP](../MCP.md).

## Where to import from

```js
// In a browser application: everything, including the editor component.
import {
  AuditWorkbook,
  addWorksheet,
  exportWorkbookToXlsx,
} from "@ledgerglass/workbook";

// In Node: the same model functions without the interface, plus file helpers that take a path.
import { readWorkbookFile, addWorksheet } from "@ledgerglass/workbook/node";
```

`@ledgerglass/workbook/calculation-node` is the same module as `@ledgerglass/workbook/node`, kept under the name the original package used.

## What to know before you use it

- **Where this API came from.** The original SDK source was lost. These functions are the recovered application's own model functions, made public under the names assigned during recovery. Each page states what a function was observed to do, not what an original specification said. Where a name suggested more than the code does, the function was left out or the page says so.
- **Workbooks are values.** Functions that change a workbook return a new workbook and leave the one you passed in unchanged, unless a page says otherwise.
- **Only cell patches are recorded.** `applyCellPatchTransaction` increases `revision` and adds to `events`. The sheet, pane, border, number-format, hyperlink, table, filter and pivot functions return a new workbook with the same revision and no new event. If your application relies on the revision number or the event log, record those changes yourself.
- **Previews guard against stale edits.** Several edits come in two steps: build a preview, then apply it. The apply step refuses a preview made from an older state of the workbook.
- **Errors are plain `Error`s with a message for the user**, except where a page names a specific class (`FileProcessingError`, `DataValidationError`, `WorkbookSpecError`).
- **Catalogue arrays are shared.** The function-name and option lists are the arrays the engine itself reads, and most are not frozen. Treat them as read-only.
