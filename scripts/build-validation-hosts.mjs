/** Rebuild all standalone validation pages after build:lib and build:portable. */
import { readdir } from "node:fs/promises";
import { execFileSync } from "node:child_process";
import { fileURLToPath } from "node:url";

const app = fileURLToPath(new URL("../", import.meta.url));
const scripts = (await readdir(new URL("./", import.meta.url)))
  .filter((name) => /^build-.*-host\.mjs$/.test(name))
  .sort();
for (const name of scripts) {
  execFileSync(process.execPath, [fileURLToPath(new URL(name, import.meta.url))], {
    cwd: app,
    stdio: "inherit",
  });
}
console.log(`Built ${scripts.length} validation hosts.`);
