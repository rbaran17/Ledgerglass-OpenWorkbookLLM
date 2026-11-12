/**
 * Bundles the page into public/. Run from the repository root after `npm run build:lib`:
 *   node examples/source-gateway/build.mjs
 */
import { mkdir, writeFile } from "node:fs/promises";
import { fileURLToPath } from "node:url";
import { build } from "esbuild";

const here = fileURLToPath(new URL(".", import.meta.url));
export async function buildPage() {
  await mkdir(`${here}public/assets`, { recursive: true });
  await build({
    entryPoints: [`${here}src/main.jsx`],
    outdir: `${here}public/assets`,
    bundle: true,
    format: "esm",
    splitting: true,
    jsx: "automatic",
    minify: true,
    define: { "process.env.NODE_ENV": '"production"' },
    logLevel: "error",
  });
  await writeFile(
    `${here}public/index.html`,
    `<!doctype html><html lang="en"><head><meta charset="UTF-8"/><title>Source gateway example</title><link rel="stylesheet" href="/assets/main.css"/><style>html,body,#root{height:100%;margin:0}</style></head><body><div id="root"></div><script type="module" src="/assets/main.js"></script></body></html>`,
  );
}
if (process.argv[1] === fileURLToPath(import.meta.url)) await buildPage();
