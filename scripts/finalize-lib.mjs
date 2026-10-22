import { readFile, writeFile, readdir } from 'node:fs/promises';
// CSS is distributed through the documented style.css package export. It must
// not leave an unresolved source-relative side-effect import in declarations.
const root = new URL('../lib/', import.meta.url);
for (const file of await readdir(root,{recursive:true})) {
  if(!file.endsWith('.d.ts')) continue;
  const declaration = new URL(file,root);
  const text = await readFile(declaration, 'utf8');
  await writeFile(declaration, text.replace(/^import ["'][^"']+\.css["'];\s*$/gm, ''));
}
