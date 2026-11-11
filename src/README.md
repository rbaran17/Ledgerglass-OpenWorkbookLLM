# Recovered source

The original TypeScript project was lost. This directory was recovered from the
compiled application in `dist/` and is now the source the build uses.

| Path | Contents |
| --- | --- |
| `main.js` | The application entry: workbook UI, core model and inlined third-party libraries |
| `assistant/` | The assistant chat panel and its Markdown renderer, loaded on demand |
| `workers/*.min.js` | Prebuilt file-parser and calculation workers, imported as text and started as Blob workers |
| `styles/` | Workbook and assistant stylesheets |

## What to expect

- The code is formatted but not original: identifiers are the minifier's, and
  there are no types or comments.
- `main.js` is a single module of about 245,000 lines. Module boundaries from
  the original project are not marked.
- Third-party libraries (React, ExcelJS, the formula parser, pdf-lib, zod and
  others) are inlined at the versions the original build used. Updating
  `package.json` does not change them yet.
- The worker bundles are opaque. They contain their own copies of the core
  model and of several libraries.

## Safety net

`tests/browser/` holds behavioural checks that run against any build. They
passed identically against the original `dist/` and against a build of this
directory. Run them before and after any change here:

```sh
npm run build:container
LEDGERGLASS_BUILD=dist-container npm run test:e2e
```

## Roadmap

1. Replace inlined third-party libraries with imports from `node_modules`, one
   library at a time, so dependency updates take effect.
2. Split `main.js` into modules along the original boundaries.
3. Rebuild the workers from shared modules instead of opaque bundles.
4. Give exported and frequently used identifiers meaningful names; add types.
5. Grow the browser checks and add unit tests around the core model.
