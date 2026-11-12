# Recovered source

The original TypeScript project was lost. This directory was recovered from the
compiled application in `dist/` and is now the source the build uses.

| Path | Contents |
| --- | --- |
| `main.js` | The application entry: workbook UI, core model and the third-party libraries not yet replaced by imports |
| `assistant/` | The assistant chat panel and its Markdown renderer, loaded on demand |
| `workers/` | File-parser and calculation worker sources, bundled at build time and started as Blob workers |
| `styles/` | Workbook and assistant stylesheets |

## What to expect

- The code is formatted but not original: identifiers are the minifier's, and
  there are no types or comments.
- `main.js` is a single module of about 119,000 lines. Module boundaries from
  the original project are not marked.
- In `main.js`, React, ReactDOM, ExcelJS, Papa Parse, zod, the formula parser
  the lucide-react icons, fflate, pdf-lib and html-to-image are imported from
  `node_modules`. In the workers, zod, ExcelJS, saxes, Papa Parse,
  fflate and the formula parser are. Updates to these take effect on the next build.
- In `assistant/`, the assistant UI libraries and others are still inlined at the versions the original build
  used. Updating `package.json` does not change them yet.
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

1. Replace the remaining inlined third-party libraries with imports from
   `node_modules`, one library at a time, so dependency updates take effect.
2. Split `main.js` into modules along the original boundaries.
3. Share the core model between `main.js` and the workers instead of keeping
   three copies.
4. Give exported and frequently used identifiers meaningful names; add types.
5. Grow the browser checks and add unit tests around the core model.
