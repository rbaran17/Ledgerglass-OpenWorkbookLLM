import { defineConfig } from "vitest/config";
import react from "@vitejs/plugin-react";
export default defineConfig({
  plugins: [react()],
  test: {
    setupFiles: ["./tests/browser-primitives.setup.ts"],
    include: ["tests/*.test.ts", "tests/*.test.tsx"],
    environment: "node",
    testTimeout: 15000,
  },
});
