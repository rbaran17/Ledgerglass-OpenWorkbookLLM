import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
export default defineConfig({
  plugins: [react()],
  build: {
    // Chevrotain 7's ESM generator contains a literal CommonJS require("../api").
    // Convert mixed modules so a consuming host can bundle the emitted SDK.
    commonjsOptions: { transformMixedEsModules: true },
    outDir: "lib",
    lib: {
      entry: "src/index.ts",
      name: "Ledgerglass",
      formats: ["es"],
      fileName: "ledgerglass",
      cssFileName: "ledgerglass",
    },
    rollupOptions: { external: [/^react(?:-dom)?(?:\/|$)/] },
    chunkSizeWarningLimit: 1800,
  },
});
