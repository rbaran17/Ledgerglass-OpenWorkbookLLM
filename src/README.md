# Recovered source

The original TypeScript project was lost. This directory was recovered from the
compiled application in `dist/` and is now the source the build uses.

| Path                                 | Contents                                                                                                     |
| ------------------------------------ | ------------------------------------------------------------------------------------------------------------ |
| `main.js`                            | The entry: loads the application modules in order and re-exports what the assistant panel needs              |
| `app/`                               | The application, split into 57 ordered modules; see [`app/README.md`](app/README.md) for what each one holds |
| `assistant/AssistantChat.js`         | The assistant chat panel, loaded on demand; recovered code, with the chat UI library imported                |
| `assistant/AssistantChatMarkdown.js` | The Markdown renderer for assistant replies; rewritten as ordinary source                                    |
| `workers/`                           | File-parser and calculation worker sources, bundled at build time and started as Blob workers                |
| `styles/`                            | Workbook and assistant stylesheets                                                                           |

## What to expect

- The code is not the original. There are no types or comments. Module-level
  names (about 2,200 functions, components and constants) were assigned after
  recovery from strings, property names and usage, so they describe what the
  code appears to do and some will be imprecise. Local variables are mostly
  still single letters, except where a destructuring pattern preserved the
  real name.
- The application lives in `app/` as 60 modules, each depending only on
  lower-numbered ones; the core model is split further into twelve files
  under `app/03-workbook-model/`. The main editor component,
  `app/56-main-worksheet-editor.js` (about 3,000 lines of state, effects and
  wiring), delegates to twenty modules under `app/56-main-worksheet-editor/`:
  `render-*.js` build parts of the interface and `*-actions.js` hold the
  editor's commands. Each takes what it reads from the component as one
  explicit object, so their long parameter lists show how entangled the
  editor's state still is. Module
  boundaries follow dependency order, not the original project's files.
- In `main.js`, React, ReactDOM, ExcelJS, Papa Parse, zod, saxes, the formula parser,
  the lucide-react icons, fflate, pdf-lib and html-to-image are imported from
  `node_modules`. In the workers, zod, ExcelJS, saxes, Papa Parse,
  fflate and the formula parser are. Updates to these take effect on the next build.
- No known third-party library is inlined any more. This rests on licence
  banners and known signatures, so a small library without a banner could
  still be present.
- The workers (`workers/*.worker.js`) are bundled by
  `scripts/worker-source-plugin.mjs`. They import the core model from
  `app/01-…`, `app/02-…`, `app/03-workbook-model/`, `app/06-…` and
  `app/06c-package-and-worksheet-parsing.js` instead of carrying their own copies: the
  calculation worker is about 100 lines (its message handler) and the file
  parser about 3,900 (parsing that only runs there, still with minified
  names). Those modules must stay free of React, DOM access and
  worker-source imports so the workers can load them; `app/01b-…` and
  `app/06b-…` exist to keep such things out of them, and `app/06c-…` holds
  the file-format code that modules 04, 05, 07 and 08 share with the worker.
  About 800 lines in the file parser still closely resemble application
  code but differ in what they call, so they are not shared. The MCP SDK is
  not part of the browser application.

## Safety net

`tests/unit/` holds unit tests for the core model and for both workers. They
run in Node (`npm run test:unit`): the core modules need no browser, and the
workers are loaded with a stand-in `self`.

`tests/browser/` holds behavioural checks that run against any build. They
passed identically against the original `dist/` and against a build of this
directory. Run them before and after any change here:

```sh
npm run build:container
LEDGERGLASS_BUILD=dist-container npm run test:e2e
```

## Roadmap

1. Reduce the coupling in `app/56-main-worksheet-editor/`: turn the render
   functions into components and the action factories into hooks that own
   their state, so each takes a handful of inputs instead of dozens.
2. Name the file-parser worker's own parsing code and reconcile the last
   near-duplicates (`importXlsxWorkbook`, `normalizeWorkbookModel`) with
   their application twins.
3. Review the assigned names, name the remaining locals, and add types.
4. Keep growing the unit tests and browser checks; XLSX export, printing
   and the data tools have the least coverage.
