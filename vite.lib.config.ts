import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import { readFileSync } from "node:fs";
import { workerSourceVite } from "./scripts/worker-source-plugin.mjs";

const { dependencies } = JSON.parse(readFileSync("package.json", "utf8"));
const external = Object.keys(dependencies);

/** Builds the embeddable library into lib/. Dependencies stay external; the host installs them. */
export default defineConfig({
  plugins: [workerSourceVite(), react()],
  build: {
    outDir: "lib",
    emptyOutDir: true,
    sourcemap: false,
    lib: {
      entry: "src/index.js",
      formats: ["es"],
      fileName: () => "ledgerglass.js",
      cssFileName: "ledgerglass",
    },
    rollupOptions: {
      external: (id) =>
        external.some((name) => id === name || id.startsWith(`${name}/`)),
    },
  },
});
