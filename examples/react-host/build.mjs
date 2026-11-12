// Builds the host page into public/: a bundle of src/main.jsx (with the library and its
// stylesheet) under public/assets/, and an index.html that loads it.
// Run from the repository root: node examples/react-host/build.mjs
import { mkdir, rm, writeFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";
import { build } from "esbuild";

const here = path.dirname(fileURLToPath(import.meta.url));
const publicDir = path.join(here, "public");

// The content security policy the page is served with. `worker-src blob:` is required: the
// editor's parser and calculation workers are started from blob URLs.
const contentSecurityPolicy = [
  "default-src 'self'",
  "script-src 'self'",
  "style-src 'self' 'unsafe-inline'",
  "img-src 'self' data: blob:",
  "worker-src blob:",
  "connect-src 'self'",
  "font-src 'self' data:",
].join("; ");

export async function buildHost() {
  await rm(publicDir, { recursive: true, force: true });
  await mkdir(path.join(publicDir, "assets"), { recursive: true });
  await build({
    entryPoints: [path.join(here, "src/main.jsx")],
    outdir: path.join(publicDir, "assets"),
    bundle: true,
    format: "esm",
    jsx: "automatic",
    minify: true,
    // React checks this to decide whether to include development warnings.
    define: { "process.env.NODE_ENV": '"production"' },
    logLevel: "warning",
  });
  await writeFile(
    path.join(publicDir, "index.html"),
    `<!doctype html>
<html lang="en">
<head>
<meta charset="UTF-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<meta http-equiv="Content-Security-Policy" content="${contentSecurityPolicy}" />
<link rel="icon" href="data:," />
<title>Ledgerglass host example</title>
<link rel="stylesheet" href="assets/main.css" />
</head>
<body>
<div id="root"></div>
<script type="module" src="assets/main.js"></script>
</body>
</html>
`,
  );
  return publicDir;
}

if (import.meta.url === pathToFileURL(process.argv[1]).href) {
  const out = await buildHost();
  console.log(`Built ${path.relative(process.cwd(), out)}/`);
}
