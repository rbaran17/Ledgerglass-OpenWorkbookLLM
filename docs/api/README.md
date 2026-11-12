# API reference

The model functions the library exports, by area. Every function listed on these pages was run before it was documented, has TypeScript declarations in `types/sdk/` and has tests in `tests/unit/sdk-*.test.js`.

| Area                             | What it covers                                                                              | Exports |
| -------------------------------- | ------------------------------------------------------------------------------------------- | ------: |
| [Files](files.md)                | Open and save XLSX, delimited text and native JSON                                          |       8 |
| [Worksheets](worksheets.md)      | Add, rename, move, hide and protect sheets; panes, merges, hidden rows; workbook properties |      25 |
| [Cells](cells.md)                | Addresses, ranges and selections; borders, number formats, hyperlinks, AutoSum              |      25 |
| [Tables](tables.md)              | Worksheet tables, filters and saved filter views, pivot tables                              |      25 |
| [Rules](rules.md)                | Data validation, conditional formatting, defined names, review threads, formula watches     |      22 |
| [Calculation](calculation.md)    | Calculation settings and evidence, the calculator object, function-name catalogues, what-if |      22 |
| [Recording](#edits-are-recorded) | Edits recorded in `revision` and `events`, `createActorEdits`, `recordWorkbookChange`       |       2 |
| [Node entry](node.md)            | Files by path and all calculated values at once, for scripts and servers                    |       3 |

Runnable examples are in [`examples/`](../../examples/README.md). [All exports](exports.md) lists every public name on one page, with the entry points that export it.

The editor component, the assistant adapters and the core functions (`createBlankWorkbook`, `normalizeWorkbookInput`, `getWorkbookCalculation`, `applyCellPatchTransaction`) are described in [Using Ledgerglass as a library](../LIBRARY.md). External data sources (the gateway, connectors and browser adapter) are in [External data sources](../SOURCES.md). The agent-rendered workbook is in [Showing a workbook from an agent](../AGENT-UI.md), and the MCP server in [MCP](../MCP.md).

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
- **Previews guard against stale edits.** Several edits come in two steps: build a preview, then apply it. The apply step refuses a preview made from an older state of the workbook.
- **Errors are plain `Error`s with a message for the user**, except where a page names a specific class (`FileProcessingError`, `DataValidationError`, `WorkbookSpecError`).
- **Catalogue arrays are shared.** The function-name and option lists are the arrays the engine itself reads, and most are not frozen. Treat them as read-only.

## Edits are recorded

Every edit made through this API advances the workbook's `revision` by one and adds one entry to `events`. The edit returns a new workbook; the one you passed in is not changed.

- `applyCellPatchTransaction` lists each changed cell in the `changes` of its event.
- Every other edit adds an event with an empty `changes` list and an `action` named after the function in words: `addWorksheet` adds "Add worksheet", `setCellHyperlink` adds "Set cell hyperlink".
- The default actor is `"Preparer"`. `createActorEdits(actor)` returns the same edit functions, recording under another actor.
- An edit that changes nothing and returns the same workbook object is not recorded. Checks, lookups and previews record nothing; applying a preview records one event.
- This matches what the editor application records for the same edits.
- `protectWorkbookStructure`, `unprotectWorkbookStructure` and `updateSheetProtection` are asynchronous; the workbook they resolve to is recorded like any other edit.
- The apply step of a previewed edit is logged by what it does: "Apply defined name change" and "Apply conditional format change"; `withCalculationSettings` is logged as "Change calculation settings".

`recordWorkbookChange(before, after, { action, actor? })` records an edit you made by hand. It returns `after` with `revision` one higher than `before`, one new event and the modified time updated. `action` is required (1 to 200 characters); `actor` defaults to `"Preparer"` (1 to 100 characters). It returns `after` unchanged when it is the same object as `before` or has already been recorded.

`createActorEdits(actor)` returns a frozen object with every edit function of the worksheets, cells, tables, rules and calculation areas, recording under `actor`. It throws a `TypeError` for a blank or over-long actor.

```js
import {
  addWorksheet,
  createActorEdits,
  createBlankWorkbook,
  recordWorkbookChange,
} from "@ledgerglass/workbook";

const book = createBlankWorkbook("Budget");
const added = addWorksheet(book, "Data");
added.revision; // 1
added.events[0].action; // "Add worksheet"

const edits = createActorEdits("A. Patel");
edits.addWorksheet(book, "Notes").events[0].actor; // "A. Patel"

const renamed = { ...book, name: "Budget FY26" };
const recorded = recordWorkbookChange(book, renamed, {
  action: "Rename workbook",
  actor: "A. Patel",
});
recorded.events[0]; // { id, timestamp, actor: "A. Patel", action: "Rename workbook", changes: [] }
```
