import { defineConfig } from "vitest/config";

/** Unit tests for the core model; they run in Node without a browser. */
export default defineConfig({
  test: {
    include: ["tests/unit/*.test.js"],
    environment: "node",
    testTimeout: 15000,
  },
});
