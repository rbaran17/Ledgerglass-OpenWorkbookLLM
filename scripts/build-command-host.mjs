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
  entryPoints: ["examples/CommandWorkbookHost.tsx"],
  outfile: "command-host.js",
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
const hostStyles = `html,body,#root{margin:0;height:100%;width:100%}body{font-family:Arial,sans-serif}.command-host{height:100%;display:flex;flex-direction:column}.command-host>header{padding:12px 16px;background:#eef4f7;color:#193945;font-size:12px}.command-host>header strong{font-size:16px}.command-host p{margin:6px 0}.command-host output{display:block;overflow-wrap:anywhere;font-family:monospace;margin-top:6px}.command-windows{display:flex;flex:1;min-height:0;gap:12px;padding:12px;background:#dbe3e9}.command-windows>section{display:flex;flex-direction:column;width:50%;min-width:0;min-height:0;border:1px solid #a8b9c4;border-radius:5px;overflow:hidden}.command-host-field{display:flex;align-items:center;gap:8px;padding:7px;background:#f8fafb;font-size:12px}.command-host input:not([type=checkbox]){padding:4px;border:1px solid #a8b9c4;border-radius:3px}.command-windows .lg-app{flex:1;min-height:0}.command-host footer{padding:7px 16px;font-size:12px}`;
const html = `<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Workbook command acceptance</title><style>${(stylesheet + hostStyles).replaceAll("</style", "<\\/style")}</style></head><body><div id="root"></div><script type="module">${script.replaceAll("</script", "<\\/script")}</script></body></html>`;
const output = new URL("dist/commands.html", root);
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
  `Created dist/commands.html from the compiled public SDK (${(Buffer.byteLength(html) / 1024 / 1024).toFixed(2)} MB).`,
);
console.log(
  "Open /commands.html in a separate tab; the fixture never writes browser storage.",
);
