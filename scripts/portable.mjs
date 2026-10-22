import { writeFile } from "node:fs/promises";
import { fileURLToPath } from "node:url";
import { build } from "esbuild";
// Bundle every lazy module into the standalone file. A Vite production build may
// contain several JS/CSS chunks, so selecting the first asset loses the app.
const result = await build({
  absWorkingDir: fileURLToPath(new URL("../", import.meta.url)),
  entryPoints: ["src/main.tsx"],
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
await writeFile(new URL("../Embedded-Workbook.html", import.meta.url), html.replace('<html lang="en">','<html lang="en" data-mode="embedded">'));
console.log(`Portable application created: Ledgerglass.html (${(Buffer.byteLength(html) / 1024 / 1024).toFixed(2)} MB).`);
