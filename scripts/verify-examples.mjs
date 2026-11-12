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

const host = await run("examples/react-host/verify.mjs");
assert.equal(
  host.code,
  0,
  `examples/react-host/verify.mjs\n${host.stdout}\n${host.stderr}`,
);
const hostChecks = host.stdout.match(/^PASS /gm)?.length ?? 0;
assert.ok(
  hostChecks >= 8,
  "the React host check ran fewer checks than expected",
);
console.log(`PASS examples/react-host (${hostChecks} browser checks)`);
console.log(`${count + 1} examples verified.`);
