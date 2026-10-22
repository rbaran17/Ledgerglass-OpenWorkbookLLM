import { build } from "esbuild";
await build({
  entryPoints: ["src/integrations/source-connectors.ts"],
  outfile: "lib/source-connectors.js",
  bundle: true,
  platform: "node",
  format: "esm",
  target: "node22",
  packages: "external",
  minify: true,
  logLevel: "warning",
});
await build({
  entryPoints: ["src/integrations/sqlite-source-worker.ts"],
  outfile: "lib/sqlite-source-worker.cjs",
  bundle: true,
  platform: "node",
  format: "cjs",
  target: "node22",
  minify: true,
  logLevel: "warning",
});
console.log("Built server-only source connectors and isolated SQLite process.");
