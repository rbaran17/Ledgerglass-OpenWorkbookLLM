/** Local-only acceptance server. Its two fixed PDF targets contain synthetic fixtures. */
import { createServer } from "node:http";
import { mkdir, readFile, writeFile } from "node:fs/promises";
const origin = "http://127.0.0.1:4178";
const destinations = {
  "/__print-evidence/letter.pdf": new URL(
    "../examples/print-output/Letter-portrait.pdf",
    import.meta.url,
  ),
  "/__print-evidence/a4.pdf": new URL(
    "../examples/print-output/A4-landscape.pdf",
    import.meta.url,
  ),
};
createServer(async (req, res) => {
  try {
    if (req.headers.host !== "127.0.0.1:4178") {
      res.writeHead(403).end();
      return;
    }
    const path = new URL(req.url ?? "/", origin).pathname;
    if (req.method === "GET" && (path === "/" || path === "/print.html")) {
      res.writeHead(200, {
        "Content-Type": "text/html; charset=utf-8",
        "Cache-Control": "no-store",
      });
      res.end(await readFile(new URL("../dist/print.html", import.meta.url)));
      return;
    }
    const destination = destinations[path];
    if (req.method !== "POST" || !destination) {
      res.writeHead(404).end();
      return;
    }
    if (
      req.headers.origin !== origin ||
      req.headers["content-type"] !== "application/pdf"
    ) {
      res.writeHead(403).end();
      return;
    }
    let size = 0;
    const chunks = [];
    for await (const chunk of req) {
      size += chunk.length;
      if (size > 50 * 1024 * 1024) {
        res.writeHead(413).end();
        return;
      }
      chunks.push(chunk);
    }
    const bytes = Buffer.concat(chunks);
    if (bytes.subarray(0, 5).toString() !== "%PDF-") {
      res.writeHead(400).end();
      return;
    }
    await mkdir(new URL("../examples/print-output/", import.meta.url), {
      recursive: true,
    });
    await writeFile(destination, bytes);
    res
      .writeHead(201, { "Content-Type": "text/plain" })
      .end("Saved synthetic verification PDF.");
  } catch (error) {
    res.writeHead(500).end("Verification file could not be saved.");
    console.error(error);
  }
}).listen(4178, "127.0.0.1", () =>
  console.log(`Print acceptance fixture: ${origin}/print.html`),
);
