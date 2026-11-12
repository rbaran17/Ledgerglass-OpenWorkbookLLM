import { expect, test } from "vitest";
import { readFileSync } from "node:fs";

/** Names a module exports, read from its source so the module does not have to load here. */
function exportedNames(path) {
  const source = readFileSync(path, "utf8");
  const names = new Set();
  for (const block of source.matchAll(/export\s*\{([^}]*)\}/g))
    for (const item of block[1].split(",")) {
      const name = item
        .trim()
        .split(/\s+as\s+/)
        .pop();
      if (name) names.add(name);
    }
  return [...names].sort();
}

function declaredNames(path) {
  const source = readFileSync(path, "utf8");
  return [...source.matchAll(/^export (?:declare )?(?:const|function) (\w+)/gm)]
    .map((match) => match[1])
    .sort();
}

test("the library's declarations cover exactly what it exports", () => {
  expect(declaredNames("types/index.d.ts")).toEqual(
    exportedNames("src/index.js"),
  );
});

test("the MCP server's declarations cover exactly what it exports", () => {
  expect(declaredNames("types/mcp.d.ts")).toEqual(
    exportedNames("src/mcp/server.js"),
  );
});
