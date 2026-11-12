import { build } from "esbuild";
import { existsSync } from "node:fs";
import { dirname, resolve } from "node:path";

const SUFFIX = "?worker-source";

// Bundles a worker entry into a self-contained classic-script string.
async function bundleWorker(file) {
  const result = await build({
    entryPoints: [file],
    bundle: true,
    write: false,
    format: "iife",
    platform: "browser",
    target: "es2022",
    minify: true,
    legalComments: "none",
    define: { "process.env.NODE_ENV": '"production"' },
    logLevel: "warning",
  });
  return result.outputFiles[0].text;
}

// Vite plugin: `import src from "./x.worker.js?worker-source"` yields the bundled worker text.
export function workerSourceVite() {
  return {
    name: "worker-source",
    enforce: "pre",
    resolveId(source, importer) {
      if (!source.endsWith(SUFFIX)) return null;
      const path = source.slice(0, -SUFFIX.length);
      let absolutePath = resolve(
        importer ? dirname(importer) : process.cwd(),
        path,
      );
      // Vitest can pass the importer relative to the project root ("/src/app/x.js").
      if (!existsSync(absolutePath) && importer?.startsWith("/"))
        absolutePath = resolve(process.cwd(), `.${dirname(importer)}`, path);
      return absolutePath + SUFFIX;
    },
    async load(id) {
      if (!id.endsWith(SUFFIX)) return null;
      let path = id.slice(0, -SUFFIX.length);
      // A cached module id can also arrive relative to the project root.
      if (!existsSync(path)) path = resolve(process.cwd(), `.${path}`);
      this.addWatchFile(path);
      const text = await bundleWorker(path);
      return `export default ${JSON.stringify(text)};`;
    },
  };
}

// esbuild plugin: the same `?worker-source` import contract for esbuild builds.
export function workerSourceEsbuild() {
  return {
    name: "worker-source",
    setup(build) {
      build.onResolve({ filter: /\?worker-source$/ }, (args) => {
        const path = args.path.slice(0, -SUFFIX.length);
        return {
          path: resolve(args.resolveDir, path),
          namespace: "worker-source",
        };
      });
      build.onLoad(
        { filter: /.*/, namespace: "worker-source" },
        async (args) => {
          const text = await bundleWorker(args.path);
          return { contents: text, loader: "text" };
        },
      );
    },
  };
}
