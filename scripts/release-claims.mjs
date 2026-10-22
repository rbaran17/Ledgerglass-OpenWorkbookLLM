import { readFile, writeFile, readdir } from "node:fs/promises";
import { resolve, relative } from "node:path";
import { pathToFileURL, fileURLToPath } from "node:url";
import { createHash } from "node:crypto";
import Papa from "papaparse";

const CONTRACT_HASH =
  "b771fd1b8042060b7fff1143dd7ed09cd52d5d17491fd9139f7632187e3cc3fd";
const ACCEPTED = "Implemented (local acceptance)";
const BASELINE_ACCEPTED = [
  "AW-001",
  "AW-002",
  "AW-004",
  "AW-005",
  "AW-007",
  "AW-008",
  "AW-013",
  "AW-014",
  "AW-017",
  "AW-018",
  "AW-019",
  "AW-021",
  "AW-022",
  "AW-023",
  "AW-024",
  "AW-025",
  "AW-026",
  "AW-027",
  "AW-049",
];
const statuses = [ACCEPTED, "Partial", "Planned"];
const fail = (message) => {
  throw Error(`Release claims: ${message}`);
};
const require = (condition, message) => {
  if (!condition) fail(message);
};
const fields = {
  "Requirement ID": "id",
  "Capability family": "family",
  Capability: "capability",
  Priority: "priority",
  "Target release": "targetRelease",
  "Acceptance criterion 1": "ac1",
  "Acceptance criterion 2": "ac2",
};
const tally = (rows, key) =>
  rows.reduce(
    (counts, row) => ((counts[row[key]] = (counts[row[key]] ?? 0) + 1), counts),
    {},
  );

/** Reporting scope only: the complete product contract and implementation stay intact. */
export function selectCoverageRows(input) {
  const scope = input.coverageScope;
  require(scope?.schemaVersion === 1 && scope.id === "base-excel", "unknown coverage scope");
  require(Array.isArray(scope.includedRequirementIds) &&
    Array.isArray(scope.excludedGroups), "coverage scope must declare included and excluded requirements");
  const excluded = scope.excludedGroups.flatMap((group) => {
    require(group.reason?.trim() && Array.isArray(group.requirementIds), "coverage exclusions need a reason and requirement IDs");
    return group.requirementIds;
  });
  const all = [...scope.includedRequirementIds, ...excluded];
  require(new Set(all).size === all.length &&
    all.length === input.matrix.length &&
    input.matrix.every((row) => all.includes(row["Requirement ID"])),
  "coverage scope must classify every retained requirement exactly once");
  require(Object.keys(scope.displayFamilies ?? {}).every((id) =>
    scope.includedRequirementIds.includes(id)), "display families must refer to included requirements");
  return scope.includedRequirementIds.map((id) => {
    const row = input.matrix.find((item) => item["Requirement ID"] === id);
    return {
      ...row,
      "Capability family": scope.displayFamilies?.[id] ?? row["Capability family"],
    };
  });
}

/** A consistency/release-evidence gate; passing it does not establish Office or production conformance. */
export function inspectReleaseClaims(input) {
  const {
    contract,
    matrix,
    requirements,
    cases,
    evidence,
    scope,
    runtime,
    files,
    testReport,
  } = input;
  require(contract.schemaVersion === 1 &&
    scope.schemaVersion === 1 &&
    evidence.schemaVersion ===
      1, "unknown contract, scope or evidence version");
  require(contract.requirements.length === 108 &&
    matrix.length === 108, "all 108 original requirements must remain present");
  require(new Set(matrix.map((r) => r["Requirement ID"])).size ===
    108, "duplicate requirement ID");
  require(Object.keys(scope.claims).sort().join("|") ===
    [
      "auditStandardCompliance",
      "externalLLMTested",
      "immutableLocalHistory",
      "productionAuthorization",
      "universalExcelParity",
    ].join("|"), "unexpected compatibility claim fields");
  require(Object.values(scope.claims).every(
    (v) => v === false,
  ), "this local release cannot claim Excel parity, audit compliance, immutable history, production authorization or actual provider acceptance");
  require(JSON.stringify(scope.capacity) ===
    JSON.stringify(
      runtime.limits,
    ), "declared capacity differs from the compiled runtime");
  require(scope.supportedFunctions === runtime.functions &&
    scope.assistantFunctions ===
      runtime.assistantFunctions, "formula capability counts differ from the compiled runtime");
  require(Array.isArray(scope.externalGates) &&
    scope.externalGates.length >= 3 &&
    scope.externalGates.every(
      (s) => typeof s === "string" && s.trim(),
    ), "external acceptance gates must remain explicit");
  const baseline = new Set(evidence.baselineAccepted);
  require(JSON.stringify([...baseline].sort()) ===
    JSON.stringify(BASELINE_ACCEPTED), "invalid previous acceptance inventory");
  require(new Set(evidence.closures.map((c) => c.id)).size ===
    evidence.closures.length, "duplicate closure record");
  const checks = new Set();
  const evidencePaths = (value) => {
    const paths = value
      .split(";")
      .map((s) => s.trim())
      .filter(Boolean);
    require(paths.length > 0, "empty implementation evidence");
    for (const p of paths) {
      require(!p.includes("..") &&
        !p.startsWith("/") &&
        files.has(p), `missing or invalid evidence path ${p}`);
      if (/^tests\/.*\.test\.tsx?$/.test(p)) checks.add(p);
    }
    return paths;
  };
  for (const original of contract.requirements) {
    const row = matrix.find((r) => r["Requirement ID"] === original.id);
    require(row, `missing ${original.id}`);
    for (const [field, key] of Object.entries(fields))
      require(row[field] ===
        original[key], `${original.id} changed original ${field}`);
    require(statuses.includes(
      row["Delivery status"],
    ), `${original.id} has an invalid delivery status`);
    const heading = `### ${original.id} — ${original.capability}\n`;
    const section = requirements.split(heading)[1]?.split("\n### ")[0];
    require(section?.includes(
      `Delivery: ${row["Delivery status"]}\n`,
    ), `${original.id} document and matrix status disagree`);
    require(section?.includes(
      `\n\n${original.description}\n- AC1: ${original.ac1}\n- AC2: ${original.ac2}`,
    ), `${original.id} description or acceptance criteria changed`);
    require(row[
      "Remaining acceptance gap"
    ]?.trim(), `${original.id} must state its remaining gap or accepted boundary`);
    if (row["Delivery status"] === "Planned") {
      require(row["Implementation evidence"] ===
        "No implementation claim", `${original.id} Planned status contradicts an implementation claim`);
      continue;
    }
    const paths = evidencePaths(row["Implementation evidence"]);
    if (row["Delivery status"] !== ACCEPTED) continue;
    require(paths.some((p) =>
      /^tests\/.*\.test\.tsx?$/.test(p),
    ), `${original.id} acceptance needs executable test evidence`);
    if (!baseline.has(original.id)) {
      const closure = evidence.closures.find((c) => c.id === original.id);
      require(closure, `${original.id} cannot be promoted without a closure record`);
      for (const key of ["description", "ac1", "ac2"]) {
        require(closure[
          key
        ]?.assessment?.trim(), `${original.id} lacks an assessment of the complete ${key}`);
        const linked = evidencePaths(closure[key]?.evidence ?? "");
        require(linked.some((p) =>
          /^tests\/.*\.test\.tsx?$/.test(p),
        ), `${original.id} ${key} lacks a test`);
      }
      require(closure.remainingBoundary?.trim(), `${original.id} must distinguish local acceptance from remaining product/host boundaries`);
    }
  }
  require(evidence.closures.every((c) =>
    matrix.some(
      (r) => r["Requirement ID"] === c.id && r["Delivery status"] === ACCEPTED,
    ),
  ), "closure evidence claims an unaccepted requirement");
  require(cases.length > 0 &&
    new Set(cases.map((c) => c.ID)).size ===
      cases.length, "missing or duplicate validation cases");
  for (const c of cases)
    require(Object.values(c).length === 10 &&
      Object.values(c).every(
        (v) => typeof v === "string" && v.trim(),
      ), `incomplete validation case ${c.ID}`);
  for (const original of contract.requirements)
    for (const n of [1, 2]) {
      const test = cases.find(
        (c) => c.ID === `REQ-${original.id.slice(3)}-AC${n}`,
      );
      require(test &&
        test["Expected result"] ===
          `AC${n}: ${original[`ac${n}`]}`, `${original.id} original AC${n} validation case changed or disappeared`);
    }
  require(cases.filter((c) => c.ID.startsWith("FN-")).length ===
    runtime.functions, "function validation cases do not match registered functions");
  if (testReport) {
    require(testReport.success &&
      testReport.numFailedTests === 0 &&
      testReport.numPendingTests ===
        0, "test run failed or contains skipped cases");
    require(scope.verification.testsPassed === testReport.numPassedTests &&
      scope.verification.testFiles ===
        testReport.testResults
          .length, "verification counts do not match the executed report");
    for (const file of checks) {
      const result = testReport.testResults.find((r) =>
        r.name.replaceAll("\\", "/").endsWith("/" + file),
      );
      require(result &&
        result.assertionResults.length > 0 &&
        result.assertionResults.every(
          (t) => t.status === "passed",
        ), `${file} lacks a complete passing execution`);
    }
  }
  const coverage = selectCoverageRows(input);
  require(JSON.stringify(input.coverageMatrix) === JSON.stringify(coverage),
    "base Excel coverage differs from its declared projection; run update:coverage");
  return {
    requirements: tally(matrix, "Delivery status"),
    coverage: {
      scope: input.coverageScope.id,
      requirements: coverage.length,
      acceptanceCriteria: coverage.length * 2,
      statuses: tally(coverage, "Delivery status"),
      excludedRequirements: matrix.length - coverage.length,
    },
    validationCases: cases.length,
    caseStatuses: tally(cases, "Status"),
    tests: scope.verification,
    checkedTestFiles: checks.size,
  };
}

export function renderCoverage(input, result) {
  const matrix = selectCoverageRows(input);
  const concise = ["Partial", "Planned"]
    .map((status) => {
      const rows = matrix
        .filter((row) => row["Delivery status"] === status)
        .sort((a, b) => a["Requirement ID"].localeCompare(b["Requirement ID"]));
      return `### ${status} — ${rows.length}\n\n${rows.map((row) => `- ${row["Requirement ID"]}: ${row.Capability}`).join("\n") || "None."}`;
    })
    .join("\n\n");
  const gapSections = ["Partial", "Planned"].map((status) => {
    const rows = matrix
      .filter((row) => row["Delivery status"] === status)
      .sort((a, b) => a["Requirement ID"].localeCompare(b["Requirement ID"]));
    const entries = rows.map((row) => {
      const entry = `- **${row["Requirement ID"]} — ${row.Capability}** (${row.Priority} · target ${row["Target release"]}): ${row["Remaining acceptance gap"]}`;
      return status === "Planned"
        ? `${entry}\n\n  **AC1:** ${row["Acceptance criterion 1"]}\n\n  **AC2:** ${row["Acceptance criterion 2"]}`
        : entry;
    });
    return `### ${status} — ${rows.length} requirements\n\n${entries.join("\n\n") || "None."}`;
  });
  const gaps = `## Remaining gaps — concise list\n\nOne entry per open requirement. Partial means a working subset still has acceptance gaps; Planned means the original criteria remain open. The detailed gaps immediately below specify what remains within each capability.\n\n${concise}\n\n## Requirement gaps\n\nEach entry lists the remaining acceptance gap from the current feature matrix, grouped by delivery status and ordered by requirement ID. Planned items include both original acceptance criteria, which remain open. Priority and target release are the original planning labels. Local acceptance boundaries remain documented in the complete feature matrix.\n\n${gapSections.join("\n\n")}\n\n`;
  const n = result.coverage.statuses;
  return `# Current requirement coverage — base Excel\n\nThis reporting view contains **${matrix.length} base Excel requirements and ${matrix.length * 2} acceptance criteria**: workbook editing, formulas, formatting, data preparation, analysis, file exchange, collaboration, security and core usability. Legacy AW identifiers remain stable for traceability.\n\n| Status | Base Excel requirements |\n|---|---:|\n| Locally accepted | ${n[ACCEPTED] ?? 0} |\n| Partial | ${n.Partial ?? 0} |\n| Planned | ${n.Planned ?? 0} |\n| **Total** | **${matrix.length}** |\n\n**Open requirements: ${(n.Partial ?? 0) + (n.Planned ?? 0)} of ${matrix.length}.** These counts reflect the requested reporting scope; status changes require local acceptance evidence and no application feature was removed. They are not percentages of Excel feature parity. The full product register retains 108 requirements, including the extensions excluded from this reporting view.\n\nAudit workflows, custom AI/MCP, embedding and application-specific requirements are excluded from this coverage sheet and its gap totals. All implemented features, original requirements and acceptance evidence are retained in the [full product register](ALL-FEATURE-REGISTER.csv). The exact inclusion and exclusion rules are recorded in [coverage scope](COVERAGE-SCOPE.json). Only reporting family labels for core security, collaboration and usability have changed; original criteria and historical evidence remain intact; current status follows recorded acceptance.\n\nThe full product validation register retains **${result.validationCases} cases**. The recorded automated run has **${result.tests.testsPassed} passing tests across ${result.tests.testFiles} files** across the complete application, including retained extensions. These are full-product counts, not base-Excel-only acceptance counts. Automated evidence, live browser observations and manual/external acceptance remain distinct.\n\nThe local implementation has declared limits: ${input.scope.capacity.maxRows.toLocaleString("en-US")} rows, ${input.scope.capacity.maxCols.toLocaleString("en-US")} columns and ${input.scope.capacity.maxCells.toLocaleString("en-US")} stored cells. Worksheet address space and populated-data processing capacity are different limits. There is no universal Excel parity or independent Microsoft Excel conformance claim.\n\n${gaps}[Base Excel coverage sheet](FEATURE-MATRIX.csv) · [Closure evidence](LOCAL-ACCEPTANCE-EVIDENCE.json) · [Verification](VERIFICATION.md) · [Full validation checklist](VALIDATION-CHECKLIST.md)\n`;
}

export async function loadReleaseClaims(root, reportPath) {
  const read = (p) => readFile(resolve(root, p), "utf8"),
    json = async (p) => JSON.parse(await read(p));
  const contractText = await read("docs/ORIGINAL-REQUIREMENT-CONTRACT.json");
  require(createHash("sha256").update(contractText).digest("hex") ===
    CONTRACT_HASH, "the original requirement contract changed");
  const parse = (p) => {
    const result = Papa.parse(p, { header: true, skipEmptyLines: true });
    require(result.errors.length === 0, "invalid CSV");
    return result.data;
  };
  const files = new Set();
  async function walk(dir) {
    for (const entry of await readdir(dir, { withFileTypes: true })) {
      if (
        [
          "node_modules",
          ".git",
          "work",
          "lib",
          "dist",
          "test-results",
          "playwright-report",
        ].includes(entry.name)
      )
        continue;
      const path = resolve(dir, entry.name);
      if (entry.isDirectory()) await walk(path);
      else files.add(relative(root, path).replaceAll("\\", "/"));
    }
  }
  await walk(root);
  const sdk = await import(
    pathToFileURL(resolve(root, "lib/ledgerglass.js")).href
  );
  return {
    contract: JSON.parse(contractText),
    matrix: parse(await read("docs/ALL-FEATURE-REGISTER.csv")),
    coverageMatrix: parse(await read("docs/FEATURE-MATRIX.csv")),
    coverageScope: await json("docs/COVERAGE-SCOPE.json"),
    requirements: await read("docs/REQUIREMENTS.md"),
    cases: parse(await read("docs/VALIDATION-TEST-CASES.csv")),
    evidence: await json("docs/LOCAL-ACCEPTANCE-EVIDENCE.json"),
    scope: await json("docs/RELEASE-SCOPE.json"),
    runtime: {
      limits: sdk.LIMITS,
      functions: sdk.SUPPORTED_FUNCTIONS.length,
      assistantFunctions: sdk.ASSISTANT_FUNCTIONS.length,
    },
    files,
    ...(reportPath
      ? { testReport: JSON.parse(await readFile(resolve(reportPath), "utf8")) }
      : {}),
  };
}

if (
  process.argv[1] &&
  import.meta.url === pathToFileURL(resolve(process.argv[1])).href
) {
  try {
    const root = fileURLToPath(new URL("..", import.meta.url)),
      args = process.argv.slice(2),
      reportIndex = args.indexOf("--report");
    const input = await loadReleaseClaims(
        root,
        reportIndex < 0 ? undefined : args[reportIndex + 1],
      );
    if (args.includes("--write")) input.coverageMatrix = selectCoverageRows(input);
    const result = inspectReleaseClaims(input),
      expected = renderCoverage(input, result),
      path = resolve(root, "docs/COVERAGE-SUMMARY.md");
    if (args.includes("--write")) {
      await writeFile(resolve(root, "docs/FEATURE-MATRIX.csv"), Papa.unparse(input.coverageMatrix, { newline: "\n" }) + "\n");
      await writeFile(path, expected);
    }
    else
      require((await readFile(path, "utf8")) ===
        expected, "coverage summary is stale; review the contracts, then run update:coverage");
    console.log(
      JSON.stringify({
        ...result,
        originalContractUnchanged: true,
        externalProviderTested: false,
        executedTestsChecked: Boolean(input.testReport),
      }),
    );
  } catch (error) {
    console.error(error.message);
    process.exitCode = 1;
  }
}
