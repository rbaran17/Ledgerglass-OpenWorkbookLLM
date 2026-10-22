import { build } from "esbuild";
import { writeFile } from "node:fs/promises";
import { fileURLToPath } from "node:url";
const root = new URL("../", import.meta.url);
const result = await build({
  absWorkingDir: fileURLToPath(root),
  entryPoints: ["examples/ExternalSourcesWorkbookHost.tsx"],
  outfile: "external-sources.js",
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
const script = result.outputFiles.find((f) => f.path.endsWith(".js")).text,
  css = result.outputFiles.find((f) => f.path.endsWith(".css")).text;
const host =
  'html,body,#root{height:100%;margin:0}.source-host{height:100%;display:flex;flex-direction:column;color:#16254b;font:14px "Open Sans",Arial,sans-serif}.source-host>header{padding:12px 18px;background:#f2f5ff;border-bottom:1px solid #ccd6fa}.source-host h1{font-size:20px;margin:0}.source-host p{margin:7px 0}.source-host select,.source-host button{min-height:30px}.source-host>section{flex:1;min-height:0}.source-host .lg-app{height:100%;min-height:0}';
await writeFile(
  new URL("dist/external-sources.html", root),
  `<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>External sources and safe refresh</title><style>${(css + host).replaceAll("</style", "<\\/style")}</style></head><body><div id="root"></div><script type="module">${script.replaceAll("</script", "<\\/script")}</script></body></html>`,
);
console.log("Created dist/external-sources.html from the compiled SDK.");
