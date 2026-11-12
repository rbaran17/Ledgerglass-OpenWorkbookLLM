// A small Node server for the host page: it serves public/ and answers the assistant's
// POST /api/propose. No framework, so the request handling is visible.
// Run from the repository root after build.mjs: node examples/react-host/server.mjs
// Set PORT to change the port (default 4310). It listens on 127.0.0.1 only.
import { createServer } from "node:http";
import { readFile, stat } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";
import { validateAssistantProposal } from "@ledgerglass/workbook/node";

const here = path.dirname(fileURLToPath(import.meta.url));
const publicDir = path.join(here, "public");
const maxBodyBytes = 256 * 1024;

const contentTypes = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".svg": "image/svg+xml",
  ".png": "image/png",
  ".woff2": "font/woff2",
};

// Column letters and row number of an A1 address, for example "B12" -> { column: "B", row: 12 }.
function parseAddress(address) {
  const match = /^([A-Z]+)(\d+)$/.exec(address);
  return match ? { column: match[1], row: Number(match[2]) } : null;
}

function noProposal(context, explanation) {
  return {
    id: `none-${context.revision}`,
    workbookId: context.workbookId,
    revision: context.revision,
    explanation,
    citations: [],
    patches: [],
  };
}

// The demonstration "assistant": it adds up the typed numbers in one column of the approved
// selection and puts the sum in the first empty cell below them. It does not call a model.
//
// A real model call goes here, in place of this function: send `prompt` and `context` to your
// provider with a key kept in this server's environment, then turn the reply into the same
// proposal shape. Authenticate the user and check they may use this workbook before that call.
// Whatever comes back is validated below, so a reply that breaks the rules is refused.
function proposeSum(context) {
  // Only typed numbers count. Formulas are skipped so a total is never added to itself.
  const numbers = context.cells.filter((cell) => typeof cell.raw === "number");
  if (numbers.length === 0)
    return noProposal(context, "The selection has no typed numbers to add up.");

  // Use the column with the most numbers; on a tie, the leftmost column.
  const byColumn = new Map();
  for (const cell of numbers) {
    const { column } = parseAddress(cell.address);
    byColumn.set(column, [...(byColumn.get(column) ?? []), cell]);
  }
  const [column, columnNumbers] = [...byColumn.entries()].sort(
    ([left, a], [right, b]) => b.length - a.length || left.localeCompare(right),
  )[0];

  // The target must be inside the approved selection (the editor refuses edits outside it),
  // so it is an empty cell in the same column, below the last number.
  const lastRow = Math.max(
    ...columnNumbers.map((c) => parseAddress(c.address).row),
  );
  const target = context.cells
    .filter((cell) => {
      const where = parseAddress(cell.address);
      return (
        where.column === column && where.row > lastRow && cell.raw === null
      );
    })
    .sort(
      (a, b) => parseAddress(a.address).row - parseAddress(b.address).row,
    )[0];
  if (!target)
    return noProposal(
      context,
      `No empty cell below the numbers in column ${column} inside the selection.`,
    );

  // Only the numbers' cells are referenced, so the formula stays inside the selection.
  const addresses = columnNumbers
    .map((cell) => cell.address)
    .sort((a, b) => parseAddress(a).row - parseAddress(b).row);
  const formula = `=SUM(${addresses.join(",")})`;
  return {
    id: `sum-${context.revision}-${target.address}`,
    workbookId: context.workbookId,
    revision: context.revision,
    explanation: `Adds the ${addresses.length} typed numbers in column ${column} (${addresses.join(", ")}) with ${formula} in ${target.address}, the first empty cell below them.`,
    citations: addresses,
    patches: [
      {
        sheetId: context.sheetId,
        address: target.address,
        cell: { raw: formula },
      },
    ],
  };
}

function readJson(request) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    let size = 0;
    request.on("data", (chunk) => {
      size += chunk.length;
      if (size > maxBodyBytes) {
        reject(new Error("The request body is too large."));
        request.destroy();
        return;
      }
      chunks.push(chunk);
    });
    request.on("end", () => {
      try {
        resolve(JSON.parse(Buffer.concat(chunks).toString("utf8")));
      } catch {
        reject(new Error("The request body is not JSON."));
      }
    });
    request.on("error", reject);
  });
}

function sendJson(response, status, body) {
  response.writeHead(status, {
    "Content-Type": "application/json; charset=utf-8",
  });
  response.end(JSON.stringify(body));
}

async function handlePropose(request, response) {
  let body;
  try {
    body = await readJson(request);
  } catch (error) {
    return sendJson(response, 400, { error: error.message });
  }
  // The adapter sends { prompt, context, request }. The context is what the editor built from
  // the approved selection; the proposal must answer that exact context.
  if (!body || typeof body.context !== "object" || !body.context)
    return sendJson(response, 400, {
      error: "Send a context from the editor.",
    });
  try {
    const proposal = proposeSum(body.context);
    // Check the proposal against the context it was made for: same workbook and revision,
    // patches and citations inside the selection, and formulas that use only allowed functions.
    return sendJson(
      response,
      200,
      validateAssistantProposal(proposal, body.context),
    );
  } catch (error) {
    return sendJson(response, 422, { error: error.message });
  }
}

async function serveFile(response, urlPath) {
  const relative =
    urlPath === "/" ? "index.html" : decodeURIComponent(urlPath.slice(1));
  const filePath = path.resolve(publicDir, relative);
  // Refuse anything outside public/.
  if (!filePath.startsWith(publicDir + path.sep)) {
    response.writeHead(403).end("Forbidden");
    return;
  }
  try {
    const info = await stat(filePath);
    if (!info.isFile()) throw new Error("not a file");
    const body = await readFile(filePath);
    response.writeHead(200, {
      "Content-Type":
        contentTypes[path.extname(filePath)] ?? "application/octet-stream",
    });
    response.end(body);
  } catch {
    response.writeHead(404, { "Content-Type": "text/plain; charset=utf-8" });
    response.end("Not found. Run node examples/react-host/build.mjs first.");
  }
}

export function startServer({
  port = Number(process.env.PORT ?? 4310),
  host = "127.0.0.1",
} = {}) {
  const server = createServer((request, response) => {
    const { pathname } = new URL(request.url, "http://localhost");
    if (request.method === "POST" && pathname === "/api/propose")
      return void handlePropose(request, response);
    if (request.method !== "GET" && request.method !== "HEAD") {
      response.writeHead(405, { Allow: "GET, HEAD, POST" }).end();
      return;
    }
    return void serveFile(response, pathname);
  });
  return new Promise((resolve, reject) => {
    server.once("error", reject);
    server.listen(port, host, () => resolve(server));
  });
}

if (import.meta.url === pathToFileURL(process.argv[1]).href) {
  const server = await startServer();
  const { port } = server.address();
  console.log(`Ledgerglass host example: http://127.0.0.1:${port}/`);
}
