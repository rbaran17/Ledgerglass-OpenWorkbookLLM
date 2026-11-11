/** Serves the built application in dist/ through the production runtime. */
import { startWebServer } from "../deploy/server.mjs";
import { fileURLToPath } from "node:url";

const runtime = await startWebServer({
  root: fileURLToPath(new URL("../dist", import.meta.url)),
  host: "127.0.0.1",
  port: Number(process.env.PORT ?? 4173),
});
console.log(`Workbook: http://127.0.0.1:${runtime.address.port}/`);
for (const signal of ["SIGINT", "SIGTERM"])
  process.once(signal, () => void runtime.close());
