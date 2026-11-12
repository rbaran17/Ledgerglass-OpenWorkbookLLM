# Recovered source

The original TypeScript project was lost. This directory was recovered from the
compiled application in `dist/` and is now the source the build uses.

| Path | Contents |
| --- | --- |
| `main.js` | The entry: loads the application modules in order and re-exports what the assistant panel needs |
| `app/` | The application, split into 57 ordered modules; see [`app/README.md`](app/README.md) for what each one holds |
| `assistant/AssistantChat.js` | The assistant chat panel, loaded on demand; recovered code, with the chat UI library imported |
| `assistant/AssistantChatMarkdown.js` | The Markdown renderer for assistant replies; rewritten as ordinary source |
| `workers/` | File-parser and calculation worker sources, bundled at build time and started as Blob workers |
| `styles/` | Workbook and assistant stylesheets |

## What to expect

- The code is not the original. There are no types or comments. Module-level
  names (about 2,200 functions, components and constants) were assigned after
  recovery from strings, property names and usage, so they describe what the
  code appears to do and some will be imprecise. Local variables are mostly
  still single letters, except where a destructuring pattern preserved the
  real name.
- The application lives in `app/` as 57 modules, each depending only on
  lower-numbered ones; the core model is split further into twelve files
  under `app/03-workbook-model/`. `app/56-main-worksheet-editor.js` is about
  7,000 lines because it is essentially one very large component. Module
  boundaries follow dependency order, not the original project's files.
- In `main.js`, React, ReactDOM, ExcelJS, Papa Parse, zod, saxes, the formula parser,
  the lucide-react icons, fflate, pdf-lib and html-to-image are imported from
  `node_modules`. In the workers, zod, ExcelJS, saxes, Papa Parse,
  fflate and the formula parser are. Updates to these take effect on the next build.
- No known third-party library is inlined any more. This rests on licence
  banners and known signatures, so a small library without a banner could
  still be present.
- The workers (`workers/*.worker.js`) are bundled by
  `scripts/worker-source-plugin.mjs` and contain their own copies of the core
  model, under different minified names from `main.js`. The MCP SDK is not
  part of the browser application.

## Safety net

`tests/browser/` holds behavioural checks that run against any build. They
passed identically against the original `dist/` and against a build of this
directory. Run them before and after any change here:

```sh
npm run build:container
LEDGERGLASS_BUILD=dist-container npm run test:e2e
```

## Roadmap

1. Break up `app/56-main-worksheet-editor.js`, a single 6,400-line component
   that needs real refactoring rather than moving code.
2. Share the core model between the application and the workers instead of
   keeping three copies.
3. Review the assigned names, name the remaining locals, and add types.
4. Grow the browser checks and add unit tests around the core model.
