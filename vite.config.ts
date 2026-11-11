import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import { workerSourceVite } from "./scripts/worker-source-plugin.mjs";
export default defineConfig({
  plugins: [workerSourceVite(), react()],
  base: "./",
  build: {
    commonjsOptions: { transformMixedEsModules: true },
    chunkSizeWarningLimit: 1800,
  },
});
