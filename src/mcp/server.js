/**
 * Model Context Protocol server for one host-approved workbook selection.
 *
 * The server is read-only apart from queuing proposals for a person to review:
 * it cannot change the workbook, and every tool is confined to the selection
 * the host supplies. The host owns the workbook, the selection and the queue.
 *
 * This module was rewritten after the original source was lost. It follows the
 * surviving contract (the `createWorkbookMcpServer` options, nine tools, and
 * the application's own assistant protocol), but the tools other than
 * `workbook_get_context` and `workbook_get_capabilities` were named afresh.
 */
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import {
  databaseFunctionNames,
  dateFunctionNames,
  getWorkbookCalculation,
  statisticalFunctionNames,
  supportedFinancialFunctionNames,
  builtinFunctionNames,
  supportedTextFunctionNames,
  workbookLimits,
} from "./model.js";
import {
  buildAssistantSelectionContext,
  nonVolatileFunctionNameSet,
  resolveAssistantSelection,
  validateAssistantFormula,
  validateAssistantProposal,
} from "../app/19b-assistant-protocol.js";

const SERVER_INFO = { name: "ledgerglass-workbook", version: "1.0.0" };

const VOLATILE_FUNCTIONS = ["RAND", "RANDBETWEEN", "NOW", "TODAY"];

/** Function names called in a formula, ignoring text inside string literals. */
function functionsIn(formula) {
  const withoutStrings = formula.replace(/"(?:[^"]|"")*"/g, '""');
  return [
    ...new Set(
      [...withoutStrings.matchAll(/([A-Z_][A-Z0-9_.]*)\s*\(/gi)].map((match) =>
        match[1].toUpperCase(),
      ),
    ),
  ];
}

/** Cell addresses referenced in a formula, without `$` markers. */
function referencesIn(formula) {
  const withoutStrings = formula.replace(/"(?:[^"]|"")*"/g, '""');
  return [
    ...new Set(
      [
        ...withoutStrings.matchAll(
          /(?<![A-Z0-9_.])\$?([A-Z]{1,3})\$?([1-9]\d*)(?![A-Z0-9_]|\s*\()/gi,
        ),
      ].map((match) => `${match[1].toUpperCase()}${match[2]}`),
    ),
  ];
}

function describeFormula(formula) {
  const functions = functionsIn(formula);
  const supported = new Set(builtinFunctionNames);
  return {
    functions,
    references: referencesIn(formula),
    unsupportedFunctions: functions.filter((name) => !supported.has(name)),
    volatileFunctions: functions.filter((name) =>
      VOLATILE_FUNCTIONS.includes(name),
    ),
    referencesOtherSheets: /!/.test(formula.replace(/"(?:[^"]|"")*"/g, '""')),
  };
}

const isFormula = (raw) => typeof raw === "string" && raw.startsWith("=");

/** Context cells carry display values, so an error arrives as its text, such as "#DIV/0!". */
const isErrorValue = (value) =>
  typeof value === "string" &&
  /^#(DIV\/0!|N\/A|NAME\?|NULL!|NUM!|REF!|VALUE!|SPILL!|CALC!|CYCLE!|ERROR!|LIMIT!|TIMEOUT!|PENDING!|BLOCKED!)/.test(
    value,
  );

/** Deterministic checks over the approved cells; findings cite their cells. */
function auditCells(cells) {
  const findings = [];
  const add = (check, severity, address, message) =>
    findings.push({ check, severity, address, message });
  const byColumn = new Map();
  for (const cell of cells) {
    const column = cell.address.replace(/\d+$/, "");
    if (!byColumn.has(column)) byColumn.set(column, []);
    byColumn.get(column).push(cell);
    if (isErrorValue(cell.value))
      add(
        "error-value",
        "high",
        cell.address,
        `Cell evaluates to ${cell.value}.`,
      );
    if (
      typeof cell.raw === "string" &&
      !isFormula(cell.raw) &&
      /^\s*-?\(?\$?\d[\d,]*(\.\d+)?\)?%?\s*$/.test(cell.raw)
    )
      add(
        "number-stored-as-text",
        "medium",
        cell.address,
        "Looks like a number but is stored as text, so sums will skip it.",
      );
    if (
      isFormula(cell.raw) &&
      /(?<![A-Z0-9_.$])\d+\.\d+|\b\d{3,}\b/.test(
        cell.raw
          .replace(/"(?:[^"]|"")*"/g, "")
          .replace(/\$?[A-Z]{1,3}\$?\d+/gi, ""),
      )
    )
      add(
        "hard-coded-number-in-formula",
        "low",
        cell.address,
        "Formula contains a typed-in number; consider referencing a cell.",
      );
  }
  for (const [column, columnCells] of byColumn) {
    const formulas = columnCells.filter((cell) => isFormula(cell.raw));
    const constants = columnCells.filter(
      (cell) => typeof cell.raw === "number",
    );
    if (formulas.length >= 3 && constants.length > 0 && constants.length <= 2)
      for (const cell of constants)
        add(
          "constant-among-formulas",
          "medium",
          cell.address,
          `Column ${column} is mostly formulas but this cell holds a typed-in number.`,
        );
  }
  return findings;
}

const asResult = (structuredContent) => ({
  content: [{ type: "text", text: JSON.stringify(structuredContent, null, 2) }],
  structuredContent,
});

const asError = (error) => ({
  isError: true,
  content: [
    {
      type: "text",
      text: error instanceof Error ? error.message : "The request failed.",
    },
  ],
});

/**
 * @param {object} host
 * @param {() => object} host.getWorkbook      the authorized workbook
 * @param {() => {sheetId: string, range: string}} host.getSelection  the approved selection
 * @param {(proposal: object) => unknown} host.queueProposal  stores a validated proposal for review
 * @param {() => unknown} [host.beforeRequest]  optional hook run before each tool call
 */
function createWorkbookMcpServer(host) {
  for (const name of ["getWorkbook", "getSelection", "queueProposal"])
    if (typeof host?.[name] !== "function")
      throw new TypeError(`createWorkbookMcpServer needs ${name}().`);

  const context = () =>
    buildAssistantSelectionContext(host.getWorkbook(), host.getSelection());
  const approvedAddresses = () =>
    new Set(resolveAssistantSelection(host.getWorkbook(), host.getSelection()));
  const guarded = (handler) => async (input) => {
    try {
      // Lets a host refresh its workbook before each request (optional).
      await host.beforeRequest?.();
      return asResult(await handler(input ?? {}));
    } catch (error) {
      return asError(error);
    }
  };
  const readOnly = { readOnlyHint: true, openWorldHint: false };

  const server = new McpServer(SERVER_INFO, {
    instructions:
      "Tools read one approved worksheet selection and can queue a proposal for a person to review. Nothing here changes the workbook.",
  });

  server.registerTool(
    "workbook_get_capabilities",
    {
      title: "Capabilities and limits",
      description:
        "What this server can and cannot do, and the limits that apply to the approved selection and to proposals.",
      inputSchema: {},
      annotations: readOnly,
    },
    guarded(() => ({
      scope:
        "One host-approved rectangular selection on one visible worksheet.",
      authority: {
        reads: "Cells inside the approved selection only.",
        writes:
          "None. Proposals are queued for human review and are never applied by this server.",
        selection: "Cannot be widened or moved by a client.",
      },
      limits: {
        maxSelectionCells: workbookLimits.maxLLMCells,
        maxProposalPatches: workbookLimits.maxPatches,
      },
      formulas: {
        supportedFunctionCount: builtinFunctionNames.length,
        proposalFunctionCount: nonVolatileFunctionNameSet.size,
        volatileFunctions: VOLATILE_FUNCTIONS,
        proposals:
          "Proposed formulas may use non-volatile supported functions and may only reference cells in the approved selection.",
      },
      customFunctions: {
        execution:
          "This server cannot install, configure or run custom functions or extensions.",
      },
      tools: [
        "workbook_get_capabilities",
        "workbook_get_context",
        "workbook_read_range",
        "workbook_inspect_formula",
        "workbook_evaluate_formula",
        "workbook_get_formula_contracts",
        "workbook_check_formula_compatibility",
        "workbook_run_audit_checks",
        "workbook_queue_proposal",
      ],
    })),
  );

  server.registerTool(
    "workbook_get_context",
    {
      title: "Approved selection",
      description:
        "The approved selection: workbook id and revision, sheet, range, and each cell's stored content and calculated value.",
      inputSchema: {},
      annotations: readOnly,
    },
    guarded(() => context()),
  );

  server.registerTool(
    "workbook_read_range",
    {
      title: "Read cells",
      description:
        "Cells of a range that lies inside the approved selection. Ranges outside it are refused.",
      inputSchema: {
        range: z
          .string()
          .min(1)
          .max(40)
          .describe("A1 or A1:B2, inside the approved selection"),
      },
      annotations: readOnly,
    },
    guarded(({ range }) => {
      const current = context();
      const wanted = resolveAssistantSelection(host.getWorkbook(), {
        sheetId: current.sheetId,
        range,
      });
      const approved = new Map(
        current.cells.map((cell) => [cell.address, cell]),
      );
      if (!wanted.every((address) => approved.has(address)))
        throw Error("That range is outside the approved selection.");
      return {
        sheetId: current.sheetId,
        range: range.toUpperCase(),
        revision: current.revision,
        cells: wanted.map((address) => approved.get(address)),
      };
    }),
  );

  server.registerTool(
    "workbook_inspect_formula",
    {
      title: "Inspect a formula",
      description:
        "Functions and cell references used by a formula, and whether it could be proposed for the approved selection.",
      inputSchema: {
        formula: z.string().min(2).max(4096).describe('Starts with "="'),
      },
      annotations: readOnly,
    },
    guarded(({ formula }) => {
      if (!formula.startsWith("=")) throw Error('A formula starts with "=".');
      const details = describeFormula(formula);
      let proposable = true;
      let reason = null;
      try {
        validateAssistantFormula(formula, approvedAddresses());
      } catch (error) {
        proposable = false;
        reason = error.message;
      }
      return { formula, ...details, proposable, reason };
    }),
  );

  server.registerTool(
    "workbook_evaluate_formula",
    {
      title: "Evaluate a formula",
      description:
        "Calculates a formula against the current workbook without storing it. The formula may only use allowed functions and cells in the approved selection.",
      inputSchema: {
        formula: z.string().min(2).max(4096).describe('Starts with "="'),
        address: z
          .string()
          .max(10)
          .optional()
          .describe(
            "Cell to evaluate as, inside the approved selection. Defaults to its first cell.",
          ),
      },
      annotations: readOnly,
    },
    guarded(({ formula, address }) => {
      if (!formula.startsWith("=")) throw Error('A formula starts with "=".');
      const current = context();
      const approved = new Set(current.cells.map((cell) => cell.address));
      const evaluationAddress = (
        address ?? current.cells[0].address
      ).toUpperCase();
      if (!approved.has(evaluationAddress))
        throw Error("That address is outside the approved selection.");
      validateAssistantFormula(formula, approved);
      const { calculator } = getWorkbookCalculation(host.getWorkbook());
      return {
        formula,
        address: evaluationAddress,
        revision: current.revision,
        result: calculator.evaluateFormula(
          current.sheetId,
          evaluationAddress,
          formula,
        ),
      };
    }),
  );

  server.registerTool(
    "workbook_get_formula_contracts",
    {
      title: "Supported functions",
      description:
        "Which worksheet functions the workbook supports, by family, and which of them a proposal may use.",
      inputSchema: {
        functions: z
          .array(z.string().max(40))
          .max(100)
          .optional()
          .describe("Function names to look up. Omit for the full list."),
      },
      annotations: readOnly,
    },
    guarded(({ functions }) => {
      const families = {
        financial: supportedFinancialFunctionNames,
        text: supportedTextFunctionNames,
        date: dateFunctionNames,
        statistical: statisticalFunctionNames,
        database: databaseFunctionNames,
      };
      const supported = new Set(builtinFunctionNames);
      const familyOf = (name) =>
        Object.keys(families).find((key) => families[key].includes(name)) ??
        "general";
      const describe = (name) => ({
        name,
        supported: supported.has(name),
        family: supported.has(name) ? familyOf(name) : null,
        volatile: VOLATILE_FUNCTIONS.includes(name),
        allowedInProposals: nonVolatileFunctionNameSet.has(name),
      });
      const names = functions?.length
        ? functions.map((name) => name.toUpperCase())
        : [...builtinFunctionNames].sort();
      return { functions: names.map(describe) };
    }),
  );

  server.registerTool(
    "workbook_check_formula_compatibility",
    {
      title: "Check formulas in the selection",
      description:
        "Lists formulas in the approved selection that use unsupported or volatile functions, reference other sheets, or currently evaluate to an error.",
      inputSchema: {},
      annotations: readOnly,
    },
    guarded(() => {
      const current = context();
      const formulas = current.cells.filter((cell) => isFormula(cell.raw));
      const issues = [];
      for (const cell of formulas) {
        const details = describeFormula(cell.raw);
        const problems = [];
        if (details.unsupportedFunctions.length)
          problems.push(
            `Unsupported functions: ${details.unsupportedFunctions.join(", ")}`,
          );
        if (details.volatileFunctions.length)
          problems.push(
            `Volatile functions: ${details.volatileFunctions.join(", ")}`,
          );
        if (details.referencesOtherSheets)
          problems.push("References another worksheet.");
        if (isErrorValue(cell.value))
          problems.push(`Evaluates to ${cell.value}.`);
        if (problems.length)
          issues.push({ address: cell.address, formula: cell.raw, problems });
      }
      return {
        range: current.range,
        revision: current.revision,
        formulaCount: formulas.length,
        issues,
      };
    }),
  );

  server.registerTool(
    "workbook_run_audit_checks",
    {
      title: "Run audit checks",
      description:
        "Deterministic checks over the approved selection: error values, numbers stored as text, typed-in numbers inside formulas, and constants in formula columns. Each finding cites its cell.",
      inputSchema: {},
      annotations: readOnly,
    },
    guarded(() => {
      const current = context();
      const findings = auditCells(current.cells);
      return {
        range: current.range,
        revision: current.revision,
        cellsChecked: current.cells.length,
        findings,
      };
    }),
  );

  server.registerTool(
    "workbook_queue_proposal",
    {
      title: "Queue a proposal for review",
      description:
        "Validates a proposed change against the approved selection and current revision, then hands it to the host's review queue. It is not applied.",
      inputSchema: {
        proposal: z
          .record(z.string(), z.unknown())
          .describe(
            "{ id, workbookId, revision, explanation, citations: [addresses], patches: [{ sheetId, address, cell: { raw } }] }",
          ),
      },
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        openWorldHint: false,
      },
    },
    guarded(async ({ proposal }) => {
      const validated = validateAssistantProposal(proposal, context());
      await host.queueProposal(validated);
      return {
        queued: true,
        id: validated.id,
        applied: false,
        note: "Queued for human review. The workbook is unchanged.",
      };
    }),
  );

  return server;
}

export { createWorkbookMcpServer };
