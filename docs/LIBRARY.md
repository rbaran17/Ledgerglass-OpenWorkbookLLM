# Using Ledgerglass as a library

`npm run build:lib` builds an ES module library into `lib/` that a React application can import: the editor component, the workbook model and the assistant protocol.

> The library was rebuilt after the original SDK source was lost. The export names below were chosen for the rebuild, and the original package's exact API was not recovered. There are no TypeScript declarations, and the package is not published to a registry (`"private": true`).

## Build

```sh
npm ci
npm run build:lib   # writes lib/ledgerglass.js, lib/ledgerglass.css and code-split chunks
npm run test:lib    # mounts the built library in a host page and drives it in a browser
```

`lib/` is generated and not committed. The library keeps its dependencies external (React, React DOM, zod, ExcelJS and the others in `package.json`), so the host application installs them and its bundler resolves them. The two workers are bundled into the library as text and started as Blob workers, so the host needs no worker configuration, but its content security policy must allow `worker-src blob:`.

## Mount the editor

```jsx
import {
  AuditWorkbook,
  createBlankWorkbook,
  localDemoAdapter,
} from "@ledgerglass/workbook";
import "@ledgerglass/workbook/style.css";

const workbook = createBlankWorkbook("Budget");

export function Workbook() {
  return (
    <div style={{ height: "80vh" }}>
      <AuditWorkbook
        initialWorkbook={workbook}
        adapter={localDemoAdapter}
        onChange={(next) => console.log("revision", next.revision)}
      />
    </div>
  );
}
```

Give the container a height: the editor fills it.

| Prop                           | Meaning                                                                               |
| ------------------------------ | ------------------------------------------------------------------------------------- |
| `initialWorkbook`              | The workbook to open. Change its `id` to switch documents. Omit for a blank workbook. |
| `onChange(workbook)`           | Called after each in-memory change. It is not a durable-save acknowledgement.         |
| `onWorkbookLoaded(workbook)`   | Called when a saved or imported workbook has been loaded.                             |
| `onSelectionChange(selection)` | Called when the selected sheet or range changes.                                      |
| `adapter`                      | The assistant adapter. Defaults to the local demonstration, which calls no model.     |
| `actor`                        | Name recorded against edits. Defaults to `"Preparer"`.                                |
| `readOnly`                     | Disables editing when `true`.                                                         |
| `persistenceKey`               | Browser storage key for saving the workbook locally. Without it nothing is stored.    |
| `workbookStore`                | Your own store, used in place of browser storage.                                     |

The component accepts further props for linked workbooks, consolidation, extensions and external sources. They are listed in the parameter list of `MainWorksheetEditor` in `src/app/56-main-worksheet-editor.js`; they work, but they are not documented here because their contracts were not recovered in full.

## Connect a model

An assistant adapter is an object with a `propose` function:

```js
const adapter = {
  name: "My assistant",
  async propose(prompt, context, signal, request) {
    // context: { workbookId, revision, sheetId, sheetName, range, cells: [{ address, raw, value }] }
    return proposal; // { id, workbookId, revision, explanation, citations, patches }
  },
};
```

The editor validates whatever comes back against the approved selection and current revision before showing it for review. `createHttpAdapter` covers the common case of asking your own server:

```js
import { createHttpAdapter } from "@ledgerglass/workbook";

const adapter = createHttpAdapter({ endpoint: "/api/workbook-assistant" });
```

It POSTs `{ prompt, context, request }` as JSON and expects a proposal as the JSON reply. Authenticate the user and authorize the workbook on that server, and keep provider keys there.

## Use the model without the interface

```js
import {
  createBlankWorkbook,
  getWorkbookCalculation,
  applyCellPatchTransaction,
} from "@ledgerglass/workbook";

const workbook = createBlankWorkbook("Model only");
const sheet = workbook.sheets[0];
const edited = applyCellPatchTransaction(
  workbook,
  [{ sheetId: sheet.id, address: "A1", cell: { raw: "=6*7" } }],
  "Script",
  "Set a formula",
);
getWorkbookCalculation(edited).calculator.getTypedValue(sheet.id, "A1"); // { type: "number", value: 42 }
```

The library entry also loads the editor and its stylesheet, so it is for browser bundles. In Node, import the model directly from `src/app/03-workbook-model.js`, as the MCP server and the unit tests do.

## Exports

| Export                                                                                                                                                                               | Kind               |
| ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ | ------------------ |
| `AuditWorkbook`                                                                                                                                                                      | React component    |
| `localDemoAdapter`, `createHttpAdapter`                                                                                                                                              | Assistant adapters |
| `createBlankWorkbook`, `normalizeWorkbookInput`, `getWorkbookCalculation`, `applyCellPatchTransaction`, `buildCanonicalWorkbookSnapshot`, `shiftFormulaReferences`, `workbookLimits` | Workbook model     |
| `buildAssistantSelectionContext`, `resolveAssistantSelection`, `validateAssistantFormula`, `validateAssistantProposal`                                                               | Assistant protocol |
