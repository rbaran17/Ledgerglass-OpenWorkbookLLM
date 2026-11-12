import { defineConfig } from "vite";
import { readFileSync } from "node:fs";
import { builtinModules } from "node:module";
import { workerSourceVite } from "./scripts/worker-source-plugin.mjs";

const { dependencies } = JSON.parse(readFileSync("package.json", "utf8"));
const external = [...Object.keys(dependencies), ...builtinModules];

/**
 * Builds the Node entry into lib/node.js: the model, calculation and files without the interface.
 * Runs after the browser library build and keeps its output (emptyOutDir is off).
 */
export default defineConfig({
  plugins: [workerSourceVite()],
  build: {
    outDir: "lib",
    emptyOutDir: false,
    sourcemap: false,
    target: "node22",
    minify: false,
    lib: {
      entry: "src/node.js",
      formats: ["es"],
      fileName: () => "node.js",
    },
    rollupOptions: {
      external: (id) =>
        id.startsWith("node:") ||
        external.some((name) => id === name || id.startsWith(`${name}/`)),
    },
  },
});
