import { build } from "esbuild";
import { writeFile } from "node:fs/promises";
import { fileURLToPath } from "node:url";
const root = new URL("../", import.meta.url);
const result = await build({
  absWorkingDir: fileURLToPath(root),
  entryPoints: ["examples/ConnectedSourcesHost.tsx"],
  outfile: "connected.js",
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
const js = result.outputFiles.find((f) => f.path.endsWith(".js")).text,
  css = result.outputFiles.find((f) => f.path.endsWith(".css")).text;
const host =
  'html,body,#root{height:100%;margin:0}.connected-host{height:100%;display:flex;flex-direction:column;font:14px "Open Sans",Arial,sans-serif;color:#16254b}.connected-host>header{padding:12px 18px;background:#f2f5ff;border-bottom:1px solid #ccd6fa}.connected-host h1{font-size:20px;margin:0}.connected-host p{margin:6px 0}.connected-host header button{min-height:32px;border:1px solid #ccd6fa;border-radius:4px;background:white;color:#16254b;padding:5px 12px}.connected-host>section{flex:1;min-height:0}.connected-host .lg-app{height:100%;min-height:0}';
await writeFile(
  new URL("dist/external-connectors.html", root),
  `<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Connected source practice</title><style>${(css + host).replaceAll("</style", "<\\/style")}</style></head><body><div id="root"></div><script type="module">${js.replaceAll("</script", "<\\/script")}</script></body></html>`,
);
console.log(
  "Built dist/external-connectors.html; run npm run demo:sources for its local backend.",
);
