import { defineConfig } from "vitest/config";
import { workerSourceVite } from "./scripts/worker-source-plugin.mjs";

/** Unit tests for the core model; they run in Node without a browser. */
export default defineConfig({
  // The file functions start the bundled parser worker, as the built library does.
  plugins: [workerSourceVite()],
  test: {
    include: ["tests/unit/*.test.js"],
    environment: "node",
    testTimeout: 15000,
  },
});
