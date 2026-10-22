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
  entryPoints: ["examples/StructureWorkbookHost.tsx"],
  outfile: "structure-host.js",
  bundle: true,
  write: false,
  minify: true,
  platform: "browser",
  format: "esm",
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
const hostStyles = `html,body,#root{margin:0;height:100%;width:100%}body{font-family:Arial,sans-serif}.structure-host{height:100%;display:flex;flex-direction:column}.structure-toolbar{padding:10px 16px;background:#eef4f7;color:#193945;border-bottom:1px solid #c6d4dd;flex:none;font-size:12px}.structure-toolbar strong{font-size:15px}.structure-toolbar p{margin:5px 0}.structure-toolbar button{padding:5px 10px;margin:5px 0}.structure-toolbar output{display:block;overflow-wrap:anywhere;font-family:monospace;max-height:90px;overflow:auto}.structure-editor{flex:1;min-height:0;min-width:0}.structure-editor>.lg-app{height:100%}`;
const html = `<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Worksheet structure review</title><style>${(stylesheet + hostStyles).replaceAll("</style", "<\\/style")}</style></head><body><div id="root"></div><script type="module">${script.replaceAll("</script", "<\\/script")}</script></body></html>`;
const output = new URL("dist/structure.html", root);
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
  `Created dist/structure.html from the compiled public SDK (${(Buffer.byteLength(html) / 1024 / 1024).toFixed(2)} MB).`,
);
console.log(
  "Open /structure.html in a separate tab; the fixture never writes browser storage.",
);
