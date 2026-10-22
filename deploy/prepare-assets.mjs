import { readdir, readFile, writeFile } from 'node:fs/promises';
import { resolve, join, relative } from 'node:path';
import { gzipSync } from 'node:zlib';
import { createHash } from 'node:crypto';

const root = resolve(process.argv[2] ?? 'dist-container');
const files = [];
async function walk(dir) {
  for (const entry of await readdir(dir, { withFileTypes: true })) {
    const file = join(dir, entry.name);
    if (entry.isSymbolicLink()) throw Error('Build output must not contain symlinks.');
    if (entry.isDirectory()) await walk(file);
    else if (entry.isFile() && !entry.name.endsWith('.gz') && entry.name !== 'asset-manifest.json') files.push(file);
  }
}
await walk(root);
if (!files.includes(join(root, 'index.html'))) throw Error('Build output is missing index.html.');
const manifest = [];
for (const file of files.sort()) {
  const bytes = await readFile(file);
  manifest.push({ path: relative(root, file).replaceAll('\\', '/'), bytes: bytes.length,
    sha256: createHash('sha256').update(bytes).digest('hex') });
  if (/\.(html|js|css|json|svg|txt)$/.test(file)) await writeFile(file + '.gz', gzipSync(bytes, { level: 9 }));
}
await writeFile(join(root, 'asset-manifest.json'), JSON.stringify({ version: 1, files: manifest }, null, 2) + '\n');
console.log(JSON.stringify({ assets: files.length, uncompressedBytes: manifest.reduce((sum, file) => sum + file.bytes, 0) }));
