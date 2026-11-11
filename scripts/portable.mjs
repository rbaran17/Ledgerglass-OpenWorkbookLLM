import { readFile, writeFile } from "node:fs/promises";
import { fileURLToPath } from "node:url";
import { build } from "esbuild";
// Bundle every lazy module into the standalone file. A Vite production build may
// contain several JS/CSS chunks, so selecting the first asset loses the app.
// Vite's "?raw" imports (the prebuilt worker sources) are plain text to esbuild.
const rawText = {
  name: "raw-text",
  setup(build) {
    build.onResolve({ filter: /\?raw$/ }, async (args) => {
      const resolved = await build.resolve(args.path.slice(0, -4), {
        kind: args.kind,
        resolveDir: args.resolveDir,
      });
      return { path: resolved.path, namespace: "raw-text" };
    });
    build.onLoad({ filter: /.*/, namespace: "raw-text" }, async (args) => ({
      contents: await readFile(args.path, "utf8"),
      loader: "text",
    }));
  },
};
const result = await build({
  absWorkingDir: fileURLToPath(new URL("../", import.meta.url)),
  entryPoints: ["src/main.js"],
  plugins: [rawText],
  outfile: "portable.js",
  bundle: true,
  splitting: false,
  write: false,
  minify: true,
  platform: "browser",
  format: "esm",
  target: "es2022",
  jsx: "automatic",
  define: { "process.env.NODE_ENV": '"production"' },
  metafile: true,
  logLevel: "warning",
});
if (Object.values(result.metafile.outputs).some(output => output.imports.length))
  throw new Error("Portable application has external module dependencies.");
const js = result.outputFiles.find(file => file.path.endsWith(".js"))?.text;
const css = result.outputFiles.find(file => file.path.endsWith(".css"))?.text;
if (!js || !css) throw new Error("Portable application is missing script or styles.");
const html = `<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Workbook</title><style>${css.replaceAll("</style", "<\\/style")}</style></head><body style="margin:0"><div id="root"></div><script type="module">${js.replaceAll("</script", "<\\/script")}</script></body></html>`;
await writeFile(new URL("../Ledgerglass.html", import.meta.url), html);
// The embedded edition is the same application in ?embedded mode; link to it rather than duplicate it.
await writeFile(new URL("../Embedded-Workbook.html", import.meta.url), `<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Workbook</title><script>location.replace("Ledgerglass.html?embedded"+(location.search?"&"+location.search.slice(1):"")+location.hash)<\/script><noscript><meta http-equiv="refresh" content="0;url=Ledgerglass.html?embedded"></noscript></head><body><p><a href="Ledgerglass.html?embedded">Open the embedded workbook</a> (keep Ledgerglass.html in the same folder).</p></body></html>\n`);
console.log(`Portable application created: Ledgerglass.html (${(Buffer.byteLength(html) / 1024 / 1024).toFixed(2)} MB).`);
