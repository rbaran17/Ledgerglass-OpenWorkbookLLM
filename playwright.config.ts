import { defineConfig } from "@playwright/test";

// Serves a built application through the production runtime. Point
// LEDGERGLASS_BUILD at another build directory to check it the same way.
const build = process.env.LEDGERGLASS_BUILD ?? "dist";
const port = Number(process.env.LEDGERGLASS_TEST_PORT ?? 4173);

export default defineConfig({
  testDir: "./tests/browser",
  fullyParallel: false,
  workers: 1,
  reporter: "list",
  timeout: 60_000,
  use: {
    baseURL: `http://127.0.0.1:${port}`,
    headless: true,
    viewport: { width: 1440, height: 1000 },
    launchOptions: process.env.CHROME_PATH
      ? { executablePath: process.env.CHROME_PATH }
      : {},
  },
  webServer: {
    command: "node deploy/server.mjs",
    env: { STATIC_ROOT: build, PORT: String(port), HOST: "127.0.0.1" },
    url: `http://127.0.0.1:${port}/readyz`,
    reuseExistingServer: false,
  },
  outputDir: "./work/browser-results",
});
