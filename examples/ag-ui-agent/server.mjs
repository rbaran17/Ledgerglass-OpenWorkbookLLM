/**
 * The example's server: the page, and an AG-UI agent endpoint at POST /agent.
 *
 *   node examples/ag-ui-agent/build.mjs
 *   node examples/ag-ui-agent/server.mjs            scripted agent, no model, no key
 *   AGENT=claude node examples/ag-ui-agent/server.mjs   Claude; needs Anthropic credentials
 *
 * Then open http://127.0.0.1:4330, select A1:B5, open the AI assistant and ask for a total.
 */
import { createServer } from "node:http";
import { readFile } from "node:fs/promises";
import { extname, join, normalize } from "node:path";
import { fileURLToPath } from "node:url";
import { createAgUiEndpoint } from "./ag-ui-endpoint.mjs";
import { scriptedAgent } from "./scripted-agent.mjs";

const here = fileURLToPath(new URL(".", import.meta.url));
const contentTypes = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
};

export async function startServer({ port = 4330, agent = "scripted" } = {}) {
  // The Claude agent, and the SDK it needs, are loaded only when asked for.
  // `agent` may also be a responder function of your own.
  const respond =
    typeof agent === "function"
      ? agent
      : agent === "claude"
        ? (await import("./claude-agent.mjs")).createClaudeAgent()
        : scriptedAgent;
  const handleRun = createAgUiEndpoint(respond);
  const server = createServer(async (request, response) => {
    try {
      const pathname = new URL(request.url, "http://localhost").pathname;
      if (pathname === "/agent" && request.method === "POST")
        return await handleRun(request, response);
      const relative = normalize(pathname === "/" ? "/index.html" : pathname);
      if (request.method !== "GET" || relative.includes("..")) {
        response.writeHead(404).end();
        return;
      }
      const file = await readFile(join(here, "public", relative));
      response.writeHead(200, {
        "Content-Type":
          contentTypes[extname(relative)] ?? "application/octet-stream",
        "Content-Security-Policy":
          "default-src 'self'; script-src 'self'; style-src 'self' 'unsafe-inline'; img-src 'self' data: blob:; font-src 'self' data:; worker-src 'self' blob:; connect-src 'self'",
      });
      response.end(file);
    } catch {
      if (!response.headersSent) response.writeHead(404);
      response.end();
    }
  });
  await new Promise((resolve) => server.listen(port, "127.0.0.1", resolve));
  return server;
}

if (process.argv[1] === fileURLToPath(import.meta.url)) {
  const agent = process.env.AGENT === "claude" ? "claude" : "scripted";
  const server = await startServer({
    port: Number(process.env.PORT ?? 4330),
    agent,
  });
  console.log(
    `Open http://127.0.0.1:${server.address().port} (${agent} agent)`,
  );
}
