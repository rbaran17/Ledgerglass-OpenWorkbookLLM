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
  entryPoints: ["examples/XlsxPreservationHost.tsx"],
  outfile: "text-host.js",
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
const hostStyles =
  "html,body,#root{height:100%;margin:0}body{font-family:Arial,sans-serif}.text-host{height:100%;display:flex;flex-direction:column}.text-host>header{padding:12px 16px;background:#edf4f6;font-size:12px;color:#193c45}.text-host p{margin:6px 0}.text-host output{display:block;margin-top:6px;font:11px monospace}.text-host>section{flex:1;min-height:0;border:10px solid #dbe3e9}.text-host .lg-app{height:100%}.text-host select{padding:4px}.resource-demo-actions{display:flex;align-items:center;flex-wrap:wrap;gap:6px}.resource-demo-actions button{background:white;border:1px solid #CCD6FA;border-radius:4px;padding:6px 9px;color:#362c63;cursor:pointer}.resource-demo-actions button:hover{background:#eeebff}";
const html = `<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>XLSX source preservation</title><style>${(stylesheet + hostStyles).replaceAll("</style", "<\\/style")}</style></head><body><div id="root"></div><script type="module">${script.replaceAll("</script", "<\\/script")}</script></body></html>`;
const output = new URL("dist/xlsx-preservation.html", root);
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
  `Created dist/xlsx-preservation.html from the compiled public SDK (${(Buffer.byteLength(html) / 1024 / 1024).toFixed(2)} MB).`,
);
console.log(
  "Open /xlsx-preservation.html in a separate tab; the fixture never writes browser storage.",
);
