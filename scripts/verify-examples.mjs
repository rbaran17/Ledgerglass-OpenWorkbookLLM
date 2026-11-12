/**
 * Runs every example in examples/ and checks it still does what its page says.
 * Node scripts are compared with their recorded output (NAME.expected.txt); the React host runs
 * its own browser check. Run `npm run build:lib` first: the examples import the built package.
 */
import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { readFile, readdir } from "node:fs/promises";
import { join } from "node:path";

/** Exit codes an example returns on purpose. */
const expectedExitCodes = { "02-check-formula-errors.mjs": 1 };

function run(script) {
  return new Promise((resolve, reject) => {
    const child = spawn(process.execPath, [script], {
      stdio: ["ignore", "pipe", "pipe"],
    });
    let stdout = "";
    let stderr = "";
    child.stdout.on("data", (chunk) => (stdout += chunk));
    child.stderr.on("data", (chunk) => (stderr += chunk));
    child.on("error", reject);
    child.on("close", (code) => resolve({ code, stdout, stderr }));
  });
}

let count = 0;
for (const folder of ["node-scripts", "mcp-client"]) {
  const directory = join("examples", folder);
  const scripts = (await readdir(directory))
    .filter((name) => /^\d\d-.*\.mjs$/.test(name))
    .sort();
  for (const name of scripts) {
    const script = join(directory, name);
    const expected = await readFile(
      script.replace(/\.mjs$/, ".expected.txt"),
      "utf8",
    );
    const result = await run(script);
    assert.equal(
      result.code,
      expectedExitCodes[name] ?? 0,
      `${script} exit code\n${result.stderr}`,
    );
    assert.equal(result.stdout, expected, `${script} output changed`);
    console.log(`PASS ${script}`);
    count++;
  }
}

for (const [folder, minimumChecks] of [
  ["react-host", 8],
  ["source-gateway", 6],
]) {
  const host = await run(`examples/${folder}/verify.mjs`);
  assert.equal(
    host.code,
    0,
    `examples/${folder}/verify.mjs\n${host.stdout}\n${host.stderr}`,
  );
  const checks = host.stdout.match(/^PASS /gm)?.length ?? 0;
  assert.ok(
    checks >= minimumChecks,
    `examples/${folder} ran fewer checks than expected`,
  );
  console.log(`PASS examples/${folder} (${checks} browser checks)`);
  count++;
}
console.log(`${count} examples verified.`);
