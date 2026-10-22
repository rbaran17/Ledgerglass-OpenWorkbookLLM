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
  entryPoints: ["examples/CollaborationHost.tsx"],
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
  "html,body,#root{height:100%;margin:0}body{font-family:Arial,sans-serif;background:#f2f3f8}.collab-host{height:100%;display:flex;flex-direction:column}.collab-host>header{background:#30204f;color:white;display:flex;align-items:center;gap:14px;padding:16px 22px}.collab-host>header>span{margin-left:auto;font-size:12px;color:#ded6f0}.clara-mark{width:28px;height:28px;display:grid;place-items:center;border:1px solid #bfb0e0;border-radius:6px;font-weight:bold}.collab-host header strong span{font-weight:400;color:#ded6f0}.collab-intro{padding:12px 20px;color:#312544;font-size:12px;border-bottom:1px solid #ccd6fa}.collab-intro>strong{font-size:18px}.collab-intro p{max-width:1100px;line-height:1.4}.collab-intro small{margin-left:12px}.collab-host button{border:1px solid #ccd6fa;border-radius:4px;background:white;padding:6px 9px;cursor:pointer;color:#30204f}.collab-editors{display:grid;grid-template-columns:1fr 1fr;gap:12px;padding:12px;flex:1;min-height:0}.collab-account{display:flex;flex-direction:column;min-width:0;border:1px solid #ccd6fa;border-radius:4px;overflow:hidden;background:white}.account-summary{padding:10px;display:flex;align-items:center;gap:8px;flex-wrap:wrap;background:#e8e4f3;color:#30204f;font-size:12px}.account-summary output{font:12px monospace}.collab-account>.lg-collaboration{flex:1;min-height:0}@media(max-width:1100px){.collab-editors{grid-template-columns:1fr;overflow:auto}.collab-account{height:650px}.collab-intro small{display:block;margin:8px 0}}";
const html = `<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Governed collaboration</title><style>${(stylesheet + hostStyles).replaceAll("</style", "<\\/style")}</style></head><body><div id="root"></div><script type="module">${script.replaceAll("</script", "<\\/script")}</script></body></html>`;
const output = new URL("dist/collaboration.html", root);
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
  `Created dist/collaboration.html from the compiled public SDK (${(Buffer.byteLength(html) / 1024 / 1024).toFixed(2)} MB).`,
);
console.log(
  "Open /collaboration.html in a separate tab; the fixture never writes browser storage.",
);
