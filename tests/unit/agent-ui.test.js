import { describe, expect, test } from "vitest";
import {
  RENDER_WORKBOOK_TOOL_NAME,
  createWorkbookToolResult,
  createWorkbookToolSubscriber,
  readRenderWorkbookCall,
  renderWorkbookTool,
} from "../../src/agent-ui/ag-ui.js";
import {
  WorkbookSpecError,
  buildWorkbookFromSpec,
  describeWorkbookAsSpec,
  summarizeWorkbookForAgent,
  workbookSpecLimits,
  workbookSpecSchema,
} from "../../src/agent-ui/workbook-spec.js";
import { getWorkbookCalculation } from "../../src/app/03-workbook-model.js";

/** The raw input of every filled cell, keyed by address, for one sheet. */
function rawCells(sheet) {
  return Object.fromEntries(
    Object.entries(sheet.cells).map(([address, cell]) => [address, cell.raw]),
  );
}

/** The calculated value of one cell, as `{ type, value }`. */
function calculated(workbook, sheetIndex, address) {
  const { calculator } = getWorkbookCalculation(workbook);
  return calculator.getTypedValue(workbook.sheets[sheetIndex].id, address);
}

/** Runs the builder on a description that must be rejected and returns the error. */
function rejectionOf(spec, options) {
  try {
    buildWorkbookFromSpec(spec, options);
  } catch (error) {
    return error;
  }
  throw new Error(
    "The description was accepted; a WorkbookSpecError was expected.",
  );
}

/** Asserts a rejection whose problems include a message matching `pattern`. */
function expectProblem(spec, pattern) {
  const error = rejectionOf(spec);
  expect(error).toBeInstanceOf(WorkbookSpecError);
  expect(error.name).toBe("WorkbookSpecError");
  expect(Array.isArray(error.problems)).toBe(true);
  expect(error.problems.some((problem) => pattern.test(problem))).toBe(true);
  return error;
}

describe("cells from rows and cells", () => {
  test("rows fill the sheet from A1", () => {
    const { workbook } = buildWorkbookFromSpec({
      sheets: [
        {
          name: "Budget",
          rows: [
            ["Item", "Qty"],
            ["Apple", 3],
          ],
        },
      ],
    });
    expect(rawCells(workbook.sheets[0])).toEqual({
      A1: "Item",
      B1: "Qty",
      A2: "Apple",
      B2: 3,
    });
  });

  test("rows start at a custom startCell, also when written in lower case", () => {
    const { workbook } = buildWorkbookFromSpec({
      sheets: [
        {
          name: "Offset",
          startCell: "c3",
          rows: [
            [1, 2],
            [3, 4],
          ],
        },
      ],
    });
    expect(rawCells(workbook.sheets[0])).toEqual({
      C3: 1,
      D3: 2,
      C4: 3,
      D4: 4,
    });
  });

  test("cells are applied after rows and override the same address", () => {
    const { workbook } = buildWorkbookFromSpec({
      sheets: [{ name: "Over", rows: [[1, 2]], cells: { A1: 9, C1: "extra" } }],
    });
    expect(rawCells(workbook.sheets[0])).toEqual({ A1: 9, B1: 2, C1: "extra" });
  });

  test("null, empty text and undefined values leave the cell empty", () => {
    const { workbook } = buildWorkbookFromSpec({
      sheets: [
        {
          name: "Blanks",
          rows: [[null, "", undefined, 5]],
          cells: { B2: null, C2: "", D2: undefined },
        },
      ],
    });
    expect(Object.keys(workbook.sheets[0].cells)).toEqual(["D1"]);
    expect(rawCells(workbook.sheets[0])).toEqual({ D1: 5 });
  });

  test("a cells entry of null removes a cell that rows set", () => {
    const { workbook } = buildWorkbookFromSpec({
      sheets: [{ name: "Removed", rows: [[1, 2]], cells: { A1: null } }],
    });
    expect(rawCells(workbook.sheets[0])).toEqual({ B1: 2 });
  });

  test("lower-case and $-style addresses are normalised", () => {
    const { workbook } = buildWorkbookFromSpec({
      sheets: [
        {
          name: "Addresses",
          rows: [[1, 2]],
          cells: { b1: 9, $C$3: 7, a$5: "x" },
        },
      ],
    });
    expect(rawCells(workbook.sheets[0])).toEqual({
      A1: 1,
      B1: 9,
      C3: 7,
      A5: "x",
    });
  });

  test("numbers, booleans and text keep their types", () => {
    const { workbook } = buildWorkbookFromSpec({
      sheets: [{ name: "Types", rows: [[42, 1.5, true, false, "text"]] }],
    });
    const cells = workbook.sheets[0].cells;
    expect(cells.A1.raw).toBe(42);
    expect(cells.B1.raw).toBe(1.5);
    expect(cells.C1.raw).toBe(true);
    expect(cells.D1.raw).toBe(false);
    expect(cells.E1.raw).toBe("text");
  });

  test("a formula in a cell calculates", () => {
    const { workbook } = buildWorkbookFromSpec({
      sheets: [
        {
          name: "Totals",
          rows: [[1], [2], [3], ["=SUM(A1:A3)"]],
        },
      ],
    });
    expect(calculated(workbook, 0, "A4")).toEqual({ type: "number", value: 6 });
  });

  test("a formula can refer to another sheet", () => {
    const { workbook } = buildWorkbookFromSpec({
      sheets: [
        { name: "Data", rows: [[1], [2], [3]] },
        {
          name: "Calc",
          cells: { A1: "=SUM(Data!A1:A3)", B1: "='Data'!A2*10" },
        },
      ],
    });
    expect(calculated(workbook, 1, "A1")).toEqual({ type: "number", value: 6 });
    expect(calculated(workbook, 1, "B1")).toEqual({
      type: "number",
      value: 20,
    });
  });
});

describe("ids and title", () => {
  test("options.id gives the workbook id and numbered sheet ids", () => {
    const { workbook } = buildWorkbookFromSpec(
      { sheets: [{ name: "One" }, { name: "Two" }] },
      { id: "call-7" },
    );
    expect(workbook.id).toBe("call-7");
    expect(workbook.sheets.map((sheet) => sheet.id)).toEqual([
      "call-7:sheet:1",
      "call-7:sheet:2",
    ]);
  });

  test("without an id, a non-empty workbook id and matching sheet ids are still made", () => {
    const { workbook } = buildWorkbookFromSpec({ sheets: [{ name: "Anon" }] });
    expect(typeof workbook.id).toBe("string");
    expect(workbook.id.length).toBeGreaterThan(0);
    expect(workbook.sheets[0].id).toBe(`${workbook.id}:sheet:1`);
    const withEmptyId = buildWorkbookFromSpec(
      { sheets: [{ name: "Anon" }] },
      { id: "" },
    );
    expect(withEmptyId.workbook.id.length).toBeGreaterThan(0);
  });

  test("title defaults to Workbook and is trimmed and capped at 120 characters", () => {
    expect(buildWorkbookFromSpec({ sheets: [{ name: "S" }] }).title).toBe(
      "Workbook",
    );
    const long = buildWorkbookFromSpec({
      title: `  ${"x".repeat(200)}  `,
      sheets: [{ name: "S" }],
    });
    expect(long.title).toHaveLength(workbookSpecLimits.maxTitleLength);
    expect(long.title).toBe("x".repeat(120));
    expect(long.workbook.name).toBe(long.title);
  });

  test("readOnly is true only when the value is exactly true", () => {
    const readOnly = (value) =>
      buildWorkbookFromSpec({ readOnly: value, sheets: [{ name: "S" }] })
        .readOnly;
    expect(readOnly(true)).toBe(true);
    expect(readOnly("true")).toBe(false);
    expect(readOnly(1)).toBe(false);
    expect(readOnly(undefined)).toBe(false);
  });
});

describe("rejected descriptions", () => {
  test("a spec that is not an object, or has no sheets, is rejected", () => {
    for (const spec of [null, "sheets", []])
      expectProblem(spec, /must be an object/);
    expectProblem({}, /at least one sheet/);
    expectProblem({ sheets: [] }, /at least one sheet/);
  });

  test("more than 20 sheets is rejected and exactly 20 is accepted", () => {
    const sheets = (count) =>
      Array.from({ length: count }, (_, index) => ({ name: `S${index}` }));
    expectProblem({ sheets: sheets(21) }, /at most 20 sheets/);
    expect(
      buildWorkbookFromSpec({ sheets: sheets(20) }).workbook.sheets,
    ).toHaveLength(20);
  });

  test("a missing sheet name, or one over 31 characters, is rejected", () => {
    expectProblem({ sheets: [{ rows: [[1]] }] }, /Sheet 1 needs a name/);
    expectProblem(
      { sheets: [{ name: "x".repeat(32) }] },
      /longer than 31 characters/,
    );
    expect(
      buildWorkbookFromSpec({ sheets: [{ name: "x".repeat(31) }] }).workbook
        .sheets[0].name,
    ).toBe("x".repeat(31));
  });

  test("sheet names may not contain : \\ / ? * [ ]", () => {
    for (const name of ["a:b", "a\\b", "a/b", "a?b", "a*b", "a[b", "a]b"])
      expectProblem({ sheets: [{ name }] }, /may not contain/);
  });

  test("duplicate sheet names are rejected, ignoring case", () => {
    expectProblem(
      { sheets: [{ name: "Sheet" }, { name: "sheet" }] },
      /"sheet" is used twice/,
    );
  });

  test("a bad startCell is rejected, including one past the edge of the sheet", () => {
    for (const startCell of ["1A", "A0", "B", "XFE1"])
      expectProblem(
        { sheets: [{ name: "S", startCell, rows: [[1]] }] },
        /startCell is not a cell address/,
      );
  });

  test("rows that are not a list, or a row that is not a list, are rejected", () => {
    expectProblem(
      { sheets: [{ name: "S", rows: "1,2" }] },
      /rows must be a list of rows/,
    );
    expectProblem(
      { sheets: [{ name: "S", rows: [[1], 5] }] },
      /row 2 must be a list/,
    );
  });

  test("a bad address in cells is rejected, and cells must be an object", () => {
    expectProblem(
      { sheets: [{ name: "S", cells: { Q: 1 } }] },
      /"Q" is not a cell address/,
    );
    expectProblem(
      { sheets: [{ name: "S", cells: [1, 2] }] },
      /cells must be an object keyed by address/,
    );
  });

  test("an object, NaN or Infinity as a value is rejected", () => {
    expectProblem(
      { sheets: [{ name: "S", rows: [[{}]] }] },
      /must be a number, text, true\/false or null/,
    );
    expectProblem(
      { sheets: [{ name: "S", rows: [[NaN]] }] },
      /is not a finite number/,
    );
    expectProblem(
      { sheets: [{ name: "S", cells: { A1: Infinity } }] },
      /is not a finite number/,
    );
  });

  test("text longer than the limit is rejected", () => {
    const tooLong = "x".repeat(workbookSpecLimits.maxTextLength + 1);
    expectProblem(
      { sheets: [{ name: "S", rows: [[tooLong]] }] },
      /longer than 8000 characters/,
    );
    expect(
      buildWorkbookFromSpec({
        sheets: [
          { name: "S", rows: [["x".repeat(workbookSpecLimits.maxTextLength)]] },
        ],
      }).workbook.sheets[0].cells.A1.raw.length,
    ).toBe(workbookSpecLimits.maxTextLength);
  });

  test("more filled cells than maxCells is rejected", () => {
    const rows = Array.from({ length: workbookSpecLimits.maxCells + 1 }, () => [
      1,
    ]);
    expectProblem(
      { sheets: [{ name: "Big", rows }] },
      /filled cells; the most that can be shown is 20000/,
    );
  });

  test("several problems are reported together in one error", () => {
    const error = rejectionOf({
      sheets: [{ name: "A:B", rows: "nope" }, { name: "a" }, { name: "A" }],
    });
    expect(error).toBeInstanceOf(WorkbookSpecError);
    expect(error.problems).toHaveLength(3);
    expect(error.problems[0]).toMatch(/may not contain/);
    expect(error.problems[1]).toMatch(/rows must be a list/);
    expect(error.problems[2]).toMatch(/"A" is used twice/);
    expect(error.message).toBe(error.problems.join(" "));
  });
});

describe("summarizeWorkbookForAgent", () => {
  test("reports the title, sheet counts and formula counts", () => {
    const { workbook } = buildWorkbookFromSpec(
      {
        title: "Quarter",
        sheets: [
          {
            name: "Plan",
            rows: [
              ["Item", "Cost"],
              ["Rent", 100],
              ["Total", "=SUM(B2:B2)"],
            ],
          },
          { name: "Empty" },
        ],
      },
      { id: "sum-counts" },
    );
    const summary = summarizeWorkbookForAgent(workbook);
    expect(summary.workbookId).toBe("sum-counts");
    expect(summary.title).toBe("Quarter");
    expect(summary.sheets).toEqual([
      { name: "Plan", filledCells: 6, formulaCount: 1 },
      { name: "Empty", filledCells: 0, formulaCount: 0 },
    ]);
    expect(summary.formulaErrorCount).toBe(0);
    expect(summary.formulaErrors).toEqual([]);
  });

  test("lists formula errors with sheet, address, formula and error", () => {
    const { workbook } = buildWorkbookFromSpec({
      sheets: [{ name: "Calc", rows: [["=1/0", "=SUM(A1:A1)", "=2+2"]] }],
    });
    const summary = summarizeWorkbookForAgent(workbook);
    expect(summary.formulaErrorCount).toBe(2);
    expect(summary.formulaErrors).toEqual([
      { sheet: "Calc", address: "A1", formula: "=1/0", error: "#DIV/0!" },
      {
        sheet: "Calc",
        address: "B1",
        formula: "=SUM(A1:A1)",
        error: "#DIV/0!",
      },
    ]);
  });

  test("maxErrors caps the listed errors while formulaErrorCount keeps the total", () => {
    const { workbook } = buildWorkbookFromSpec({
      sheets: [
        {
          name: "S",
          rows: [
            ["=1/0", "=1/0", "=1/0"],
            ["=1/0", 1, "=2+2"],
          ],
        },
        { name: "T", cells: { A1: "=1/0" } },
      ],
    });
    const capped = summarizeWorkbookForAgent(workbook, { maxErrors: 2 });
    expect(capped.formulaErrorCount).toBe(5);
    expect(
      capped.formulaErrors.map((error) => `${error.sheet}!${error.address}`),
    ).toEqual(["S!A1", "S!B1"]);
    expect(
      summarizeWorkbookForAgent(workbook, { maxErrors: 0 }).formulaErrors,
    ).toEqual([]);
  });
});

describe("describeWorkbookAsSpec", () => {
  test("build, describe and build again gives the same cells", () => {
    const first = buildWorkbookFromSpec(
      {
        title: "Round",
        sheets: [
          {
            name: "A",
            rows: [
              ["h", "v"],
              ["x", 3],
            ],
            cells: { B4: "=SUM(B2:B2)", C1: true },
          },
          { name: "B", rows: [["ok"]] },
        ],
      },
      { id: "round-1" },
    );
    const description = describeWorkbookAsSpec(first.workbook);
    expect(description.title).toBe("Round");
    expect(description.readOnly).toBeUndefined();

    const second = buildWorkbookFromSpec(description, { id: "round-2" });
    expect(second.workbook.sheets.map((sheet) => sheet.name)).toEqual([
      "A",
      "B",
    ]);
    expect(second.workbook.sheets.map(rawCells)).toEqual(
      first.workbook.sheets.map(rawCells),
    );
  });
});

describe("AG-UI tool", () => {
  test("the tool definition has the name, a description and the schema as parameters", () => {
    expect(RENDER_WORKBOOK_TOOL_NAME).toBe("render_workbook");
    expect(renderWorkbookTool.name).toBe("render_workbook");
    expect(renderWorkbookTool.description.length).toBeGreaterThan(0);
    expect(renderWorkbookTool.parameters).toBe(workbookSpecSchema);
    expect(renderWorkbookTool.parameters.required).toEqual(["sheets"]);
  });

  test("readRenderWorkbookCall accepts a JSON string or an object with the same result", () => {
    const spec = { sheets: [{ name: "S", rows: [[1, "=A1*2"]] }] };
    const fromText = readRenderWorkbookCall(JSON.stringify(spec), {
      toolCallId: "t1",
    });
    const fromObject = readRenderWorkbookCall(spec, { toolCallId: "t1" });
    expect(fromText.ok).toBe(true);
    expect(fromObject.ok).toBe(true);
    expect(rawCells(fromText.workbook.sheets[0])).toEqual(
      rawCells(fromObject.workbook.sheets[0]),
    );
    expect(fromText.workbook.id).toBe("t1");
  });

  test("readRenderWorkbookCall reports invalid JSON and invalid specs as problems", () => {
    expect(readRenderWorkbookCall('{"sheets":', { toolCallId: "t2" })).toEqual({
      ok: false,
      problems: ["The tool arguments are not valid JSON."],
    });
    expect(
      readRenderWorkbookCall('{"sheets":[]}', { toolCallId: "t2" }),
    ).toEqual({
      ok: false,
      problems: ["The workbook description needs at least one sheet."],
    });
  });

  test("createWorkbookToolResult answers a good call with the summary and honours messageId", () => {
    const result = createWorkbookToolResult(
      "call-9",
      { sheets: [{ name: "S", rows: [["=1/0"]] }] },
      { messageId: "msg-1" },
    );
    expect(Object.keys(result).sort()).toEqual([
      "content",
      "id",
      "role",
      "toolCallId",
    ]);
    expect(result.id).toBe("msg-1");
    expect(result.role).toBe("tool");
    expect(result.toolCallId).toBe("call-9");
    expect(result.error).toBeUndefined();
    expect(JSON.parse(result.content)).toEqual({
      shown: true,
      workbookId: "call-9",
      title: "Workbook",
      sheets: [{ name: "S", filledCells: 1, formulaCount: 1 }],
      formulaErrorCount: 1,
      formulaErrors: [
        { sheet: "S", address: "A1", formula: "=1/0", error: "#DIV/0!" },
      ],
    });
  });

  test("createWorkbookToolResult answers a bad call with error set and shown false", () => {
    const result = createWorkbookToolResult("call-10", "{bad json", {
      messageId: "msg-2",
    });
    expect(result.id).toBe("msg-2");
    expect(result.toolCallId).toBe("call-10");
    expect(result.error).toBe("The tool arguments are not valid JSON.");
    expect(JSON.parse(result.content)).toEqual({
      shown: false,
      problems: ["The tool arguments are not valid JSON."],
    });
  });

  test("the subscriber ignores other tools and reports good and bad render calls", () => {
    const good = [];
    const bad = [];
    const subscriber = createWorkbookToolSubscriber({
      onWorkbook: (details) => good.push(details),
      onInvalid: (details) => bad.push(details),
    });

    subscriber.onToolCallEndEvent({
      event: { toolCallId: "other-1" },
      toolCallName: "something_else",
      toolCallArgs: { sheets: [{ name: "S" }] },
    });
    expect(good).toEqual([]);
    expect(bad).toEqual([]);

    subscriber.onToolCallEndEvent({
      event: { toolCallId: "c1" },
      toolCallName: RENDER_WORKBOOK_TOOL_NAME,
      toolCallArgs: { sheets: [{ name: "S", rows: [[1]] }] },
    });
    expect(good).toHaveLength(1);
    expect(good[0].toolCallId).toBe("c1");
    expect(good[0].spec).toEqual({ sheets: [{ name: "S", rows: [[1]] }] });
    expect(good[0].workbook.id).toBe("c1");
    expect(good[0].readOnly).toBe(false);
    expect(good[0].result.toolCallId).toBe("c1");
    expect(JSON.parse(good[0].result.content).shown).toBe(true);

    subscriber.onToolCallEndEvent({
      event: { toolCallId: "c2" },
      toolCallName: RENDER_WORKBOOK_TOOL_NAME,
      toolCallArgs: { sheets: [] },
    });
    expect(bad).toEqual([
      {
        toolCallId: "c2",
        problems: ["The workbook description needs at least one sheet."],
        result: expect.objectContaining({
          toolCallId: "c2",
          error: expect.any(String),
        }),
      },
    ]);
    expect(good).toHaveLength(1);
  });
});
