import { expect, test } from "vitest";
import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";

/**
 * Names a module exports, read from its source so the module does not have to load here.
 * `export { a, b }` blocks and inline `export function a` / `export const a` declarations are read,
 * and `export * from "./x.js"` is followed into that file.
 */
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
  for (const declaration of source.matchAll(
    /^export (?:async )?(?:const|function|class|let) (\w+)/gm,
  ))
    names.add(declaration[1]);
  for (const star of source.matchAll(/export \* from "(\.[^"]+)"/g))
    for (const name of exportedNames(join(dirname(path), star[1])))
      names.add(name);
  return [...names].sort();
}

/** Value names a declaration file declares, following `export *` and `export { a } from`. */
function declaredNames(path) {
  const source = readFileSync(path, "utf8");
  const names = new Set(
    [...source.matchAll(/^export (?:declare )?(?:const|function) (\w+)/gm)].map(
      (match) => match[1],
    ),
  );
  for (const block of source.matchAll(/^export \{([^}]*)\} from/gm))
    for (const item of block[1].split(",")) {
      const name = item
        .trim()
        .split(/\s+as\s+/)
        .pop();
      if (name) names.add(name);
    }
  for (const star of source.matchAll(/^export \* from "(\.[^"]+)"/gm))
    for (const name of declaredNames(join(dirname(path), `${star[1]}.d.ts`)))
      names.add(name);
  return [...names].sort();
}

test("the library's declarations cover exactly what it exports", () => {
  expect(declaredNames("types/index.d.ts")).toEqual(
    exportedNames("src/index.js"),
  );
});

test("the Node entry's declarations cover exactly what it exports", () => {
  expect(declaredNames("types/node.d.ts")).toEqual(
    exportedNames("src/node.js"),
  );
});

test("each SDK area declares exactly what it exports", () => {
  for (const area of [
    "calculation",
    "cells",
    "files",
    "rules",
    "tables",
    "worksheets",
  ])
    expect(declaredNames(`types/sdk/${area}.d.ts`), area).toEqual(
      exportedNames(`src/sdk/${area}.js`),
    );
});

test("the MCP server's declarations cover exactly what it exports", () => {
  expect(declaredNames("types/mcp.d.ts")).toEqual(
    exportedNames("src/mcp/server.js"),
  );
});

test("the source gateway's declarations cover exactly what it exports", () => {
  expect(declaredNames("types/source-gateway.d.ts")).toEqual(
    exportedNames("src/sources/gateway.js"),
  );
});

test("the source connectors' declarations cover exactly what they export", () => {
  expect(declaredNames("types/source-connectors.d.ts")).toEqual(
    exportedNames("src/sources/connectors.js"),
  );
});
