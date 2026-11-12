/**
 * Writes docs/api/exports.md: every public export in one alphabetical table, with where it is
 * imported from and the first sentence of its declaration comment.
 *
 *   node scripts/build-api-index.mjs          writes the file
 *   node scripts/build-api-index.mjs --check  fails if the file is out of date
 */
import { readFileSync, writeFileSync } from "node:fs";

const areaPages = {
  calculation: ["Calculation", "calculation.md"],
  cells: ["Cells", "cells.md"],
  files: ["Files", "files.md"],
  rules: ["Rules", "rules.md"],
  tables: ["Tables", "tables.md"],
  worksheets: ["Worksheets", "worksheets.md"],
};
const pageByName = new Map([
  ["AuditWorkbook", ["Library guide", "../LIBRARY.md"]],
  ["localDemoAdapter", ["Library guide", "../LIBRARY.md"]],
  ["createHttpAdapter", ["Library guide", "../LIBRARY.md"]],
  ["readWorkbookFile", ["Node entry", "node.md"]],
  ["writeWorkbookFile", ["Node entry", "node.md"]],
  ["calculateWorkbookValues", ["Node entry", "node.md"]],
]);
const agentNames =
  /^(AgentWorkbook|WorkbookSpecError|buildWorkbookFromSpec|describeWorkbookAsSpec|summarizeWorkbookForAgent|workbookSpec\w+|RENDER_WORKBOOK_TOOL_NAME|createWorkbookTool\w+|readRenderWorkbookCall|renderWorkbookTool|WORKBOOK_\w+|createWorkbookCatalog|workbookCatalogInstructions)$/;

/** `name -> { kind, summary }` for the value exports declared directly in a declaration file. */
function readDeclarations(path) {
  const source = readFileSync(path, "utf8");
  const found = new Map();
  const pattern =
    /(?:\/\*\*((?:(?!\*\/)[\s\S])*)\*\/\s*)?^export (?:declare )?(const|function) (\w+)/gm;
  for (const match of source.matchAll(pattern)) {
    const comment = (match[1] ?? "")
      .split("\n")
      .map((line) => line.replace(/^\s*\*?\s?/, ""))
      .join(" ")
      .replace(/\s+/g, " ")
      .trim();
    const sentence = comment.match(/^.*?[.!?](?=\s|$)/)?.[0] ?? comment;
    if (!found.has(match[3]))
      found.set(match[3], {
        kind: match[2] === "function" ? "function" : "value",
        summary: sentence.replace(/\|/g, "\\|"),
      });
  }
  return found;
}
/** The value names an entry's declaration file exports, following re-exports. */
function entryNames(path) {
  const source = readFileSync(path, "utf8");
  const names = new Set(readDeclarations(path).keys());
  for (const block of source.matchAll(/^export \{([^}]*)\} from/gm))
    for (const item of block[1].split(","))
      if (item.trim())
        names.add(
          item
            .trim()
            .split(/\s+as\s+/)
            .pop(),
        );
  for (const star of source.matchAll(/^export \* from "\.\/sdk\/(\w+)"/gm))
    for (const name of readDeclarations(`types/sdk/${star[1]}.d.ts`).keys())
      names.add(name);
  return names;
}

const declarations = new Map();
const areaByName = new Map();
for (const [area] of Object.entries(areaPages))
  for (const [name, entry] of readDeclarations(`types/sdk/${area}.d.ts`)) {
    declarations.set(name, entry);
    areaByName.set(name, area);
  }
for (const path of ["types/index.d.ts", "types/node.d.ts", "types/mcp.d.ts"])
  for (const [name, entry] of readDeclarations(path))
    if (!declarations.has(name)) declarations.set(name, entry);

const browserNames = entryNames("types/index.d.ts");
const nodeNames = entryNames("types/node.d.ts");
const mcpNames = entryNames("types/mcp.d.ts");

function pageFor(name) {
  if (areaByName.has(name)) return areaPages[areaByName.get(name)];
  if (pageByName.has(name)) return pageByName.get(name);
  if (mcpNames.has(name) && !browserNames.has(name))
    return ["MCP", "../MCP.md"];
  if (agentNames.test(name)) return ["Agent guide", "../AGENT-UI.md"];
  return ["Library guide", "../LIBRARY.md"];
}

const allNames = [
  ...new Set([...browserNames, ...nodeNames, ...mcpNames]),
].sort((left, right) =>
  left.localeCompare(right, "en", { sensitivity: "base" }),
);
const rows = allNames.map((name) => {
  const entry = declarations.get(name) ?? { kind: "value", summary: "" };
  const [title, link] = pageFor(name);
  const from = [
    browserNames.has(name) ? "main" : "",
    nodeNames.has(name) ? "node" : "",
    mcpNames.has(name) ? "mcp" : "",
  ]
    .filter(Boolean)
    .join(", ");
  return `| \`${name}\` | ${entry.kind} | ${from} | [${title}](${link}) | ${entry.summary} |`;
});

const text = `# All exports

Every public export, in alphabetical order. This page is generated from the type declarations by \`node scripts/build-api-index.mjs\`; do not edit it by hand.

"From" says which entry points export the name: **main** is \`@ledgerglass/workbook\`, **node** is \`@ledgerglass/workbook/node\` and **mcp** is \`@ledgerglass/workbook/mcp\`. ${browserNames.size} names are exported from main, ${nodeNames.size} from node and ${mcpNames.size} from mcp.

| Export | Kind | From | Described in | Summary |
| --- | --- | --- | --- | --- |
${rows.join("\n")}
`;

const target = "docs/api/exports.md";
if (process.argv.includes("--check")) {
  const { format } = await import("prettier");
  const formatted = await format(text, { parser: "markdown" });
  if (readFileSync(target, "utf8") !== formatted) {
    console.error(
      `${target} is out of date. Run: node scripts/build-api-index.mjs`,
    );
    process.exit(1);
  }
  console.log(`${target} is current (${allNames.length} exports).`);
} else {
  const { format } = await import("prettier");
  writeFileSync(target, await format(text, { parser: "markdown" }));
  console.log(`Wrote ${target} (${allNames.length} exports).`);
}
