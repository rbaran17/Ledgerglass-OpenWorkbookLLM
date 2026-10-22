/** Bundle the independent acceptance host against the already built public SDK. */
import { build } from "esbuild";
import { access, readFile, writeFile } from "node:fs/promises";
import { fileURLToPath } from "node:url";

const root = new URL("../", import.meta.url);
try {
  await Promise.all([
    access(new URL("lib/ledgerglass.js", root)),
    access(new URL("dist/index.html", root)),
  ]);
} catch {
  throw new Error(
    "Build the public SDK and browser application first: npm run build:lib && npm run build:portable",
  );
}
const bundle = await build({
  absWorkingDir: fileURLToPath(root),
  entryPoints: ["examples/ConcurrentWorkbookHost.tsx"],
  outfile: "concurrency-host.js",
  bundle: true,
  write: false,
  minify: true,
  platform: "browser",
  target: "es2022",
  jsx: "automatic",
  define: { "process.env.NODE_ENV": '"production"' },
  logLevel: "warning",
});
const script = bundle.outputFiles.find((file) =>
  file.path.endsWith(".js"),
)?.text;
const stylesheet = bundle.outputFiles.find((file) =>
  file.path.endsWith(".css"),
)?.text;
if (!script || !stylesheet)
  throw new Error(
    "The public SDK bundle must include JavaScript and workbook styles.",
  );
const hostStyles = `html,body,#root{margin:0;height:100%;width:100%}body{font-family:Arial,sans-serif}.concurrent-host{height:100%;display:flex;flex-direction:column}.concurrent-toolbar{padding:10px 14px;background:#eef4f7;color:#193945;border-bottom:1px solid #c6d4dd;flex:none;font-size:13px}.concurrent-toolbar strong{font-size:15px}.concurrent-context{margin-left:18px}.concurrent-controls{display:flex;align-items:center;gap:8px;flex-wrap:wrap;margin-top:8px}.concurrent-controls a{margin-right:10px;color:#125c8b}.concurrent-controls button{border:1px solid #aebfcb;border-radius:4px;background:white;color:#183744;padding:6px 10px;cursor:pointer}.concurrent-controls button:disabled{opacity:.5;cursor:default}.concurrent-controls button:focus-visible,.concurrent-controls a:focus-visible{outline:3px solid #258abd;outline-offset:2px}.concurrent-instructions{margin:8px 0 6px;color:#415d6c;font-size:12px}.concurrent-status{display:flex;gap:20px;align-items:baseline;flex-wrap:wrap;font-size:12px}.concurrent-status output{color:#224f39}.concurrent-editor{flex:1;min-height:0;min-width:0}.concurrent-editor>.lg-app{height:100%}`;
const html = `<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Concurrent workbook acceptance host</title><style>${(stylesheet + hostStyles).replaceAll("</style", "<\\/style")}</style></head><body><div id="root"></div><script type="module">${script.replaceAll("</script", "<\\/script")}</script></body></html>`;
const output = new URL("dist/concurrency.html", root);
await writeFile(output, html);
// Keep the result independently reviewable: it has no external asset requests.
const written = await readFile(output, "utf8");
if (
  /<(?:script[^>]*\bsrc|link[^>]*\bhref)\s*=/i.test(
    written.slice(0, written.indexOf('<script type="module">')),
  )
)
  throw new Error(
    "The acceptance host unexpectedly references external assets.",
  );
console.log(
  `Created dist/concurrency.html from the compiled public SDK (${(Buffer.byteLength(html) / 1024 / 1024).toFixed(2)} MB).`,
);
console.log(
  "Open /concurrency.html?run=review-01&document=shared in two same-origin browser tabs.",
);
