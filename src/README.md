# Recovered source

The original TypeScript project was lost. This directory was recovered from the
compiled application in `dist/` and is now the source the build uses.

| Path | Contents |
| --- | --- |
| `main.js` | The application entry: workbook UI, core model and the third-party libraries not yet replaced by imports |
| `assistant/AssistantChat.js` | The assistant chat panel, loaded on demand; recovered code, with the chat UI library imported |
| `assistant/AssistantChatMarkdown.js` | The Markdown renderer for assistant replies; rewritten as ordinary source |
| `workers/` | File-parser and calculation worker sources, bundled at build time and started as Blob workers |
| `styles/` | Workbook and assistant stylesheets |

## What to expect

- The code is formatted but not original: identifiers are the minifier's, and
  there are no types or comments.
- `main.js` is a single module of about 117,500 lines. Module boundaries from
  the original project are not marked.
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

2. Split `main.js` into modules along the original boundaries.
3. Share the core model between `main.js` and the workers instead of keeping
   three copies.
4. Give exported and frequently used identifiers meaningful names; add types.
5. Grow the browser checks and add unit tests around the core model.
