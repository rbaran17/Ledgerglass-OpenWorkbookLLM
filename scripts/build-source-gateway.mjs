import { build } from "esbuild";
await build({
  entryPoints: ["src/integrations/source-gateway.ts"],
  outfile: "lib/source-gateway.js",
  bundle: true,
  platform: "node",
  format: "esm",
  target: "node22",
  minify: true,
  logLevel: "warning",
});
console.log("Built the server-only external-source gateway.");
