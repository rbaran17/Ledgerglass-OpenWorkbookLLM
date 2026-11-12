/**
 * The workbook description an agent sends when it wants a spreadsheet shown to the user, and the
 * conversion from that description to a real workbook. Shared by the AG-UI tool and the A2UI
 * component. Loads in plain Node as well as the browser.
 *
 * Everything an agent sends is treated as untrusted data: it is checked against fixed limits and
 * then passed through the same normalisation as an imported file.
 */
import { workbookLimits } from "../app/02-text-and-financial-functions.js";
import {
  createBlankWorkbook,
  getWorkbookCalculation,
  normalizeWorkbookInput,
} from "../app/03-workbook-model.js";

/** Limits for agent-supplied workbooks; smaller than the editor's own so a reply stays light. */
export const workbookSpecLimits = Object.freeze({
  maxSheets: 20,
  maxCells: 20000,
  maxTextLength: 8000,
  maxTitleLength: 120,
  maxSheetNameLength: 31,
});

const cellValueSchema = {
  type: ["string", "number", "boolean", "null"],
  description:
    'A cell: number, text, true/false, or null for empty. Text starting with "=" is a formula, for example "=SUM(B2:B5)".',
};

/** JSON Schema of the description. Used as the AG-UI tool parameters and in the A2UI catalog. */
export const workbookSpecSchema = Object.freeze({
  type: "object",
  properties: {
    title: {
      type: "string",
      description: "Short name of the workbook, shown above the grid.",
    },
    sheets: {
      type: "array",
      minItems: 1,
      maxItems: workbookSpecLimits.maxSheets,
      description: "The worksheets, in tab order.",
      items: {
        type: "object",
        properties: {
          name: {
            type: "string",
            description:
              "Tab name, unique in the workbook, at most 31 characters, without : \\ / ? * [ ]",
          },
          rows: {
            type: "array",
            description:
              "Cell values row by row, starting at startCell. The first row is usually the header.",
            items: { type: "array", items: cellValueSchema },
          },
          startCell: {
            type: "string",
            description: 'Top-left cell of rows. Defaults to "A1".',
          },
          cells: {
            type: "object",
            description:
              'Individual cells by address, for example {"D7": "=SUM(D2:D6)"}. Applied after rows.',
            additionalProperties: cellValueSchema,
          },
        },
        required: ["name"],
        additionalProperties: false,
      },
    },
    readOnly: {
      type: "boolean",
      description:
        "True to show the workbook without allowing edits. Defaults to false.",
    },
  },
  required: ["sheets"],
  additionalProperties: false,
});

/** Thrown for a description that cannot be shown. `problems` lists every reason in plain words. */
export class WorkbookSpecError extends Error {
  constructor(problems) {
    super(problems.join(" "));
    this.name = "WorkbookSpecError";
    this.problems = problems;
  }
}

const addressPattern = /^\$?([A-Za-z]{1,3})\$?([1-9][0-9]{0,6})$/;
const forbiddenSheetNameCharacters = /[:\\/?*[\]]/;

function columnNumberFromLetters(letters) {
  let column = 0;
  for (const letter of letters.toUpperCase())
    column = column * 26 + (letter.charCodeAt(0) - 64);
  return column;
}

function columnLettersFromNumber(column) {
  let letters = "";
  for (let rest = column; rest > 0; rest = Math.floor((rest - 1) / 26))
    letters = String.fromCharCode(65 + ((rest - 1) % 26)) + letters;
  return letters;
}

function parseAddress(text) {
  const match = typeof text === "string" ? addressPattern.exec(text) : null;
  if (!match) return null;
  const column = columnNumberFromLetters(match[1]);
  const row = Number(match[2]);
  if (column > workbookLimits.maxCols || row > workbookLimits.maxRows)
    return null;
  return { row, column };
}

function readCellValue(value, where, problems) {
  if (value === null || value === undefined || value === "") return undefined;
  if (typeof value === "number") {
    if (Number.isFinite(value)) return value;
    problems.push(`${where} is not a finite number.`);
    return undefined;
  }
  if (typeof value === "boolean") return value;
  if (typeof value === "string") {
    if (value.length <= workbookSpecLimits.maxTextLength) return value;
    problems.push(
      `${where} is longer than ${workbookSpecLimits.maxTextLength} characters.`,
    );
    return undefined;
  }
  problems.push(`${where} must be a number, text, true/false or null.`);
  return undefined;
}

function readSheet(sheetSpec, index, usedNames, problems) {
  const label = `Sheet ${index + 1}`;
  const cells = {};
  if (!sheetSpec || typeof sheetSpec !== "object" || Array.isArray(sheetSpec)) {
    problems.push(`${label} must be an object.`);
    return { name: label, cells };
  }
  let name = typeof sheetSpec.name === "string" ? sheetSpec.name.trim() : "";
  if (!name) problems.push(`${label} needs a name.`);
  else if (name.length > workbookSpecLimits.maxSheetNameLength)
    problems.push(`${label}: the name is longer than 31 characters.`);
  else if (forbiddenSheetNameCharacters.test(name))
    problems.push(`${label}: the name may not contain : \\ / ? * [ ].`);
  else if (usedNames.has(name.toLowerCase()))
    problems.push(`${label}: the name "${name}" is used twice.`);
  usedNames.add(name.toLowerCase());
  name ||= label;

  if (sheetSpec.rows !== undefined) {
    const origin = parseAddress(sheetSpec.startCell ?? "A1");
    if (!origin) problems.push(`${label}: startCell is not a cell address.`);
    else if (!Array.isArray(sheetSpec.rows))
      problems.push(`${label}: rows must be a list of rows.`);
    else
      sheetSpec.rows.forEach((rowValues, rowOffset) => {
        if (rowValues === null || rowValues === undefined) return;
        if (!Array.isArray(rowValues)) {
          problems.push(`${label}: row ${rowOffset + 1} must be a list.`);
          return;
        }
        rowValues.forEach((value, columnOffset) => {
          const row = origin.row + rowOffset;
          const column = origin.column + columnOffset;
          if (row > workbookLimits.maxRows || column > workbookLimits.maxCols) {
            problems.push(`${label}: rows run past the edge of the sheet.`);
            return;
          }
          const address = columnLettersFromNumber(column) + row;
          const raw = readCellValue(value, `${label} ${address}`, problems);
          if (raw !== undefined) cells[address] = { raw };
        });
      });
  }
  if (sheetSpec.cells !== undefined) {
    if (
      !sheetSpec.cells ||
      typeof sheetSpec.cells !== "object" ||
      Array.isArray(sheetSpec.cells)
    )
      problems.push(`${label}: cells must be an object keyed by address.`);
    else
      for (const [key, value] of Object.entries(sheetSpec.cells)) {
        const position = parseAddress(key);
        if (!position) {
          problems.push(`${label}: "${key}" is not a cell address.`);
          continue;
        }
        const address = columnLettersFromNumber(position.column) + position.row;
        const raw = readCellValue(value, `${label} ${address}`, problems);
        if (raw === undefined) delete cells[address];
        else cells[address] = { raw };
      }
  }
  return { name, cells };
}

/**
 * Builds a workbook from an agent's description.
 *
 * @param spec  The description (see `workbookSpecSchema`).
 * @param options.id  Identifier for the workbook. Pass a stable value, such as the tool call or
 *   surface identifier, so that showing the same description again keeps the same document.
 * @returns `{ workbook, title, readOnly }`
 * @throws WorkbookSpecError listing every problem found.
 */
export function buildWorkbookFromSpec(spec, options = {}) {
  const problems = [];
  if (!spec || typeof spec !== "object" || Array.isArray(spec))
    throw new WorkbookSpecError([
      "The workbook description must be an object.",
    ]);
  const sheetSpecs = spec.sheets;
  if (!Array.isArray(sheetSpecs) || sheetSpecs.length === 0)
    throw new WorkbookSpecError([
      "The workbook description needs at least one sheet.",
    ]);
  if (sheetSpecs.length > workbookSpecLimits.maxSheets)
    throw new WorkbookSpecError([
      `A workbook may have at most ${workbookSpecLimits.maxSheets} sheets.`,
    ]);
  let title = typeof spec.title === "string" ? spec.title.trim() : "";
  if (spec.title !== undefined && typeof spec.title !== "string")
    problems.push("The title must be text.");
  title = title.slice(0, workbookSpecLimits.maxTitleLength) || "Workbook";

  const usedNames = new Set();
  const sheets = sheetSpecs.map((sheetSpec, index) =>
    readSheet(sheetSpec, index, usedNames, problems),
  );
  const cellCount = sheets.reduce(
    (total, sheet) => total + Object.keys(sheet.cells).length,
    0,
  );
  if (cellCount > workbookSpecLimits.maxCells)
    problems.push(
      `The workbook has ${cellCount} filled cells; the most that can be shown is ${workbookSpecLimits.maxCells}.`,
    );
  if (problems.length) throw new WorkbookSpecError(problems.slice(0, 20));

  const blank = createBlankWorkbook(title);
  const template = blank.sheets[0];
  const idBase =
    typeof options.id === "string" && options.id ? options.id : blank.id;
  const draft = {
    ...blank,
    id: idBase,
    sheets: sheets.map((sheet, index) => ({
      ...template,
      id: `${idBase}:sheet:${index + 1}`,
      name: sheet.name,
      cells: sheet.cells,
    })),
  };
  let workbook;
  try {
    workbook = normalizeWorkbookInput(draft);
  } catch (error) {
    throw new WorkbookSpecError([
      error instanceof Error ? error.message : String(error),
    ]);
  }
  return { workbook, title, readOnly: spec.readOnly === true };
}

/**
 * A compact account of a workbook for the agent: what was shown and which formulas produced
 * errors, so the agent can correct them in its next turn.
 */
export function summarizeWorkbookForAgent(workbook, options = {}) {
  const maxErrors = options.maxErrors ?? 25;
  const { calculator } = getWorkbookCalculation(workbook);
  const formulaErrors = [];
  let errorCount = 0;
  const sheets = workbook.sheets.map((sheet) => {
    const addresses = Object.keys(sheet.cells);
    let formulaCount = 0;
    for (const address of addresses) {
      const raw = sheet.cells[address]?.raw;
      if (typeof raw !== "string" || !raw.startsWith("=")) continue;
      formulaCount++;
      const result = calculator.getTypedValue(sheet.id, address);
      if (result?.type !== "error") continue;
      errorCount++;
      if (formulaErrors.length < maxErrors)
        formulaErrors.push({
          sheet: sheet.name,
          address,
          formula: raw,
          error: String(result.value),
        });
    }
    return { name: sheet.name, filledCells: addresses.length, formulaCount };
  });
  return {
    workbookId: workbook.id,
    title: workbook.name,
    sheets,
    formulaErrorCount: errorCount,
    formulaErrors,
  };
}

/** Turns a workbook back into a description, for example to report the user's edits to an agent. */
export function describeWorkbookAsSpec(workbook, options = {}) {
  return {
    title: workbook.name,
    sheets: workbook.sheets.map((sheet) => ({
      name: sheet.name,
      cells: Object.fromEntries(
        Object.entries(sheet.cells)
          .filter(([, cell]) => cell?.raw !== null && cell?.raw !== undefined)
          .map(([address, cell]) => [address, cell.raw]),
      ),
    })),
    ...(options.readOnly ? { readOnly: true } : {}),
  };
}
