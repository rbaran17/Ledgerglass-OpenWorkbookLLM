import { describe, expect, test } from "vitest";
import {
  createActorEdits,
  recordWorkbookChange,
} from "../../src/sdk/recording.js";
import * as worksheets from "../../src/sdk/worksheets.js";
import * as cells from "../../src/sdk/cells.js";
import * as tables from "../../src/sdk/tables.js";
import * as rules from "../../src/sdk/rules.js";
import * as calculation from "../../src/sdk/calculation.js";
import * as worksheetsModel from "../../src/sdk/model/worksheets.js";
import * as cellsModel from "../../src/sdk/model/cells.js";
import * as tablesModel from "../../src/sdk/model/tables.js";
import * as rulesModel from "../../src/sdk/model/rules.js";
import * as calculationModel from "../../src/sdk/model/calculation.js";
import {
  applyCellPatchTransaction,
  createBlankWorkbook,
  normalizeWorkbookInput,
} from "../../src/app/03-workbook-model.js";

const sheetIdOf = (workbook) => workbook.sheets[0].id;

/** A blank workbook whose first sheet holds the sales table: header row and three records. */
function salesBook() {
  const workbook = createBlankWorkbook("Sales");
  [
    ["Region", "Units", "Price"],
    ["North", 10, 2.5],
    ["South", 4, 3],
    ["North", 7, 1],
  ].forEach((row, rowIndex) =>
    row.forEach((raw, columnIndex) => {
      workbook.sheets[0].cells[`${"ABC"[columnIndex]}${rowIndex + 1}`] = {
        raw,
      };
    }),
  );
  return workbook;
}

function cellPatch(workbook, address, raw) {
  return { sheetId: sheetIdOf(workbook), address, cell: { raw } };
}

describe("recordWorkbookChange", () => {
  test("bumps the revision by one and leaves the before workbook as it was", () => {
    const before = createBlankWorkbook("Book");
    const after = { ...before, name: "Renamed" };
    const recorded = recordWorkbookChange(before, after, {
      action: "Rename workbook",
    });
    expect(recorded.revision).toBe(before.revision + 1);
    expect(before.revision).toBe(0);
    expect(before.events).toEqual([]);
    expect(recorded.name).toBe("Renamed");
  });

  test("appends exactly one event with the action, the actor and an empty change list", () => {
    const before = createBlankWorkbook("Book");
    const recorded = recordWorkbookChange(
      before,
      { ...before },
      { action: "Hand edit", actor: "A. Patel" },
    );
    expect(recorded.events).toHaveLength(1);
    expect(recorded.events[0]).toMatchObject({
      action: "Hand edit",
      actor: "A. Patel",
      changes: [],
    });
    expect(typeof recorded.events[0].id).toBe("string");
    expect(Number.isNaN(Date.parse(recorded.events[0].timestamp))).toBe(false);
  });

  test("uses Preparer as the default actor", () => {
    const before = createBlankWorkbook("Book");
    const recorded = recordWorkbookChange(
      before,
      { ...before },
      {
        action: "Hand edit",
      },
    );
    expect(recorded.events.at(-1).actor).toBe("Preparer");
  });

  test("updates the modified time, which is not earlier than before", () => {
    const before = createBlankWorkbook("Book");
    const recorded = recordWorkbookChange(
      before,
      { ...before },
      {
        action: "Hand edit",
      },
    );
    expect(
      Date.parse(recorded.metadata.modifiedAt) >=
        Date.parse(before.metadata.modifiedAt),
    ).toBe(true);
  });

  test("returns the after workbook unchanged when it is the same object as before", () => {
    const before = createBlankWorkbook("Book");
    expect(recordWorkbookChange(before, before, { action: "Nothing" })).toBe(
      before,
    );
  });

  test("returns the after workbook unchanged when it has already been recorded", () => {
    const before = createBlankWorkbook("Book");
    const once = recordWorkbookChange(
      before,
      { ...before },
      {
        action: "Once",
      },
    );
    expect(recordWorkbookChange(before, once, { action: "Again" })).toBe(once);
    expect(once.events).toHaveLength(1);
  });

  test("accepts an action of exactly 200 characters and refuses a missing or blank one", () => {
    const before = createBlankWorkbook("Book");
    const after = { ...before };
    expect(
      recordWorkbookChange(before, after, {
        action: "a".repeat(200),
      }).events.at(-1).action,
    ).toHaveLength(200);
    expect(() => recordWorkbookChange(before, after)).toThrow(TypeError);
    expect(() =>
      recordWorkbookChange(before, after, { action: "   " }),
    ).toThrow(TypeError);
  });

  test("refuses an action of 201 characters", () => {
    const before = createBlankWorkbook("Book");
    expect(() =>
      recordWorkbookChange(before, { ...before }, { action: "a".repeat(201) }),
    ).toThrow(TypeError);
  });

  test("refuses a blank actor and an actor of 101 characters", () => {
    const before = createBlankWorkbook("Book");
    expect(() =>
      recordWorkbookChange(
        before,
        { ...before },
        {
          action: "Edit",
          actor: " ",
        },
      ),
    ).toThrow(TypeError);
    expect(() =>
      recordWorkbookChange(
        before,
        { ...before },
        {
          action: "Edit",
          actor: "a".repeat(101),
        },
      ),
    ).toThrow(TypeError);
  });

  test("refuses arguments that are not workbooks", () => {
    const workbook = createBlankWorkbook("Book");
    expect(() =>
      recordWorkbookChange(undefined, workbook, { action: "Edit" }),
    ).toThrow(TypeError);
    expect(() =>
      recordWorkbookChange(workbook, { sheets: [] }, { action: "Edit" }),
    ).toThrow(TypeError);
    expect(() => recordWorkbookChange(null, null, { action: "Edit" })).toThrow(
      TypeError,
    );
  });

  test("the recorded workbook passes normalizeWorkbookInput with its revision and events intact", () => {
    const before = createBlankWorkbook("Book");
    const recorded = recordWorkbookChange(
      before,
      { ...before },
      {
        action: "Hand edit",
      },
    );
    const normalized = normalizeWorkbookInput(recorded);
    expect(normalized.revision).toBe(1);
    expect(normalized.events).toEqual(recorded.events);
  });
});

describe("one representative edit from each area", () => {
  test("worksheets: addWorksheet records Add worksheet", () => {
    const workbook = createBlankWorkbook("Book");
    const next = worksheets.addWorksheet(workbook, "Data");
    expect(next.revision).toBe(1);
    expect(next.events).toHaveLength(1);
    expect(next.events[0]).toMatchObject({
      action: "Add worksheet",
      actor: "Preparer",
      changes: [],
    });
    expect(workbook.events).toEqual([]);
  });

  test("worksheets: renameWorksheet records Rename worksheet", () => {
    const workbook = createBlankWorkbook("Book");
    const next = worksheets.renameWorksheet(
      workbook,
      sheetIdOf(workbook),
      "Inputs",
    );
    expect(next.sheets[0].name).toBe("Inputs");
    expect(next.events.at(-1)).toMatchObject({
      action: "Rename worksheet",
      changes: [],
    });
  });

  test("worksheets: setFreezePanesAtCell records Set freeze panes at cell", () => {
    const workbook = createBlankWorkbook("Book");
    const next = worksheets.setFreezePanesAtCell(
      workbook,
      sheetIdOf(workbook),
      "B2",
    );
    expect(next.revision).toBe(1);
    expect(next.events.at(-1).action).toBe("Set freeze panes at cell");
  });

  test("cells: applyNumberFormatTransaction records Apply number format transaction", () => {
    const workbook = salesBook();
    const next = cells.applyNumberFormatTransaction(
      workbook,
      sheetIdOf(workbook),
      "B2",
      { category: "number", decimals: 0 },
    );
    expect(next.revision).toBe(1);
    expect(next.events.at(-1)).toMatchObject({
      action: "Apply number format transaction",
      actor: "Preparer",
      changes: [],
    });
  });

  test("cells: setCellHyperlink records Set cell hyperlink", () => {
    const workbook = salesBook();
    const next = cells.setCellHyperlink(workbook, sheetIdOf(workbook), "A1", {
      kind: "external",
      url: "https://example.com",
    });
    expect(next.events.at(-1).action).toBe("Set cell hyperlink");
  });

  test("tables: createWorksheetTable records Create worksheet table", () => {
    const workbook = salesBook();
    const next = tables.createWorksheetTable(workbook, sheetIdOf(workbook), {
      name: "Sales",
      range: "A1:C4",
      headers: ["Region", "Units", "Price"],
    });
    expect(next.sheets[0].tables).toHaveLength(1);
    expect(next.revision).toBe(1);
    expect(next.events.at(-1)).toMatchObject({
      action: "Create worksheet table",
      actor: "Preparer",
      changes: [],
    });
  });

  test("tables: applyTableOperationPreview records Apply table operation preview", () => {
    const start = salesBook();
    const workbook = tables.createWorksheetTable(start, sheetIdOf(start), {
      name: "Sales",
      range: "A1:C4",
      headers: ["Region", "Units", "Price"],
    });
    const tableId = workbook.sheets[0].tables[0].id;
    const plan = tables.applyTableOperation(
      workbook,
      sheetIdOf(workbook),
      tableId,
      {
        kind: "appendRows",
        count: 1,
      },
    );
    const next = tables.applyTableOperationPreview(workbook, plan);
    expect(next.revision).toBe(workbook.revision + 1);
    expect(next.events.at(-1).action).toBe("Apply table operation preview");
  });

  test("rules: assertConditionalFormatPreviewCurrent records its change", () => {
    const workbook = createBlankWorkbook("Book");
    const preview = rules.updateConditionalFormatRules(
      workbook,
      sheetIdOf(workbook),
      [
        {
          id: "r1",
          range: "A1:A4",
          kind: "greater",
          value: "10",
          second: "",
          fill: "#ffcccc",
          color: "#000000",
        },
      ],
    );
    const next = rules.assertConditionalFormatPreviewCurrent(workbook, preview);
    expect(next.revision).toBe(1);
    expect(next.events.at(-1)).toMatchObject({
      action: "Apply conditional format change",
      changes: [],
    });
  });

  test("rules: a defined-name preview applied with assertNamePreviewCurrent records its change", () => {
    const workbook = createBlankWorkbook("Book");
    const preview = rules.applyDefinedNameOperation(workbook, {
      kind: "create",
      definition: { name: "Rate", refersTo: "=Sheet1!$A$1" },
    });
    const next = rules.assertNamePreviewCurrent(workbook, preview);
    expect(next.revision).toBe(1);
    expect(next.events.at(-1).action).toBe("Apply defined name change");
  });

  test("calculation: withCalculationSettings records Change calculation settings", () => {
    const workbook = createBlankWorkbook("Book");
    const next = calculation.withCalculationSettings(workbook, {
      ...calculation.defaultCalculationSettings,
      mode: "manual",
    });
    expect(next.calculation.mode).toBe("manual");
    expect(next.revision).toBe(1);
    expect(next.events.at(-1)).toMatchObject({
      action: "Change calculation settings",
      actor: "Preparer",
      changes: [],
    });
  });
});

describe("chained edits", () => {
  test("three edits give revisions 1, 2 and 3 and three events in order", () => {
    const start = createBlankWorkbook("Book");
    const added = worksheets.addWorksheet(start, "Data");
    const renamed = worksheets.renameWorksheet(
      added,
      added.sheets[1].id,
      "Figures",
    );
    const linked = cells.setCellHyperlink(renamed, renamed.sheets[1].id, "A1", {
      kind: "external",
      url: "https://example.com",
    });
    expect([added.revision, renamed.revision, linked.revision]).toEqual([
      1, 2, 3,
    ]);
    expect(linked.events.map((event) => event.action)).toEqual([
      "Add worksheet",
      "Rename worksheet",
      "Set cell hyperlink",
    ]);
    expect(start.revision).toBe(0);
  });
});

describe("non-edits are not recorded", () => {
  test("findSheetOrThrow returns the sheet and records nothing", () => {
    const workbook = salesBook();
    const sheet = worksheets.findSheetOrThrow(workbook, sheetIdOf(workbook));
    expect(sheet).toBe(workbook.sheets[0]);
    expect(workbook.revision).toBe(0);
    expect(workbook.events).toEqual([]);
  });

  test("summarizeWorksheets returns a summary and records nothing", () => {
    const workbook = salesBook();
    const summary = worksheets.summarizeWorksheets(workbook);
    expect(summary).not.toHaveProperty("revision");
    expect(workbook.revision).toBe(0);
    expect(workbook.events).toEqual([]);
  });

  test("refreshCalculationReport returns a calculator and report, not a workbook", () => {
    const workbook = salesBook();
    const result = calculation.refreshCalculationReport(workbook);
    expect(result).toHaveProperty("calculator");
    expect(result).toHaveProperty("report");
    expect(result).not.toHaveProperty("revision");
    expect(workbook.revision).toBe(0);
  });

  test("constants and error classes are the same values as the model's", () => {
    expect(rules.DataValidationError).toBe(rulesModel.DataValidationError);
    expect(calculation.defaultCalculationSettings).toBe(
      calculationModel.defaultCalculationSettings,
    );
    for (const [area, model] of [
      [worksheets, worksheetsModel],
      [cells, cellsModel],
      [tables, tablesModel],
      [rules, rulesModel],
      [calculation, calculationModel],
    ]) {
      for (const [name, value] of Object.entries(area)) {
        if (typeof value !== "function") expect(value, name).toBe(model[name]);
      }
    }
  });

  test("DataValidationError still works with instanceof", () => {
    const error = new rules.DataValidationError(
      createBlankWorkbook("Book"),
      [],
    );
    expect(error).toBeInstanceOf(rules.DataValidationError);
    expect(error).toBeInstanceOf(rulesModel.DataValidationError);
    expect(error).toBeInstanceOf(Error);
  });
});

describe("failing edits", () => {
  test("an edit that throws throws the same error and records nothing", () => {
    const workbook = createBlankWorkbook("Book");
    expect(() => worksheets.addWorksheet(workbook, "sheet1")).toThrow(
      "A worksheet with that name already exists.",
    );
    expect(workbook.revision).toBe(0);
    expect(workbook.events).toEqual([]);
  });
});

describe("cell patches with SDK edits", () => {
  test("a cell patch followed by an SDK edit gives revision 2 and two events", () => {
    const start = createBlankWorkbook("Book");
    const patched = applyCellPatchTransaction(start, [
      cellPatch(start, "A1", "7"),
    ]);
    const next = worksheets.renameWorksheet(
      patched,
      sheetIdOf(patched),
      "Inputs",
    );
    expect(next.revision).toBe(2);
    expect(next.events).toHaveLength(2);
    expect(next.events[0].changes.length).toBeGreaterThan(0);
    expect(next.events[0].changes[0]).toBeDefined();
    expect(next.events[1]).toMatchObject({
      action: "Rename worksheet",
      changes: [],
    });
  });
});

describe("createActorEdits", () => {
  test("events carry the given actor", () => {
    const edits = createActorEdits("A. Patel");
    const next = edits.addWorksheet(createBlankWorkbook("Book"), "Data");
    expect(next.events.at(-1)).toMatchObject({
      actor: "A. Patel",
      action: "Add worksheet",
      changes: [],
    });
  });

  test("the returned object is frozen", () => {
    const edits = createActorEdits("A. Patel");
    expect(Object.isFrozen(edits)).toBe(true);
    expect(() => {
      edits.addWorksheet = () => null;
    }).toThrow(TypeError);
  });

  test("refuses a blank or over-long actor", () => {
    expect(() => createActorEdits("")).toThrow(TypeError);
    expect(() => createActorEdits("   ")).toThrow(TypeError);
    expect(() => createActorEdits("a".repeat(101))).toThrow(TypeError);
    expect(() => createActorEdits("a".repeat(100))).not.toThrow();
  });

  test("contains the edit functions of all five areas", () => {
    const edits = createActorEdits("A. Patel");
    for (const name of [
      "addWorksheet", // worksheets
      "setCellHyperlink", // cells
      "createWorksheetTable", // tables
      "assertConditionalFormatPreviewCurrent", // rules
      "withCalculationSettings", // calculation
    ]) {
      expect(typeof edits[name], name).toBe("function");
    }
  });

  test("the module-level exports still record as Preparer", () => {
    const next = worksheets.addWorksheet(createBlankWorkbook("Book"), "Data");
    expect(next.events.at(-1).actor).toBe("Preparer");
  });
});

describe("wrapped function names", () => {
  test("wrapped functions keep their names", () => {
    expect(worksheets.addWorksheet.name).toBe("addWorksheet");
    expect(cells.applyBorderEditTransaction.name).toBe(
      "applyBorderEditTransaction",
    );
    expect(createActorEdits("A. Patel").addWorksheet.name).toBe("addWorksheet");
  });
});

describe("asynchronous edits", () => {
  test("protecting the workbook structure resolves to a recorded workbook", async () => {
    const workbook = createBlankWorkbook("Async");
    const next = await worksheets.protectWorkbookStructure(workbook, "secret");
    expect(next.revision).toBe(1);
    expect(next.events.at(-1)).toMatchObject({
      actor: "Preparer",
      action: "Protect workbook structure",
      changes: [],
    });
    expect(workbook.revision).toBe(0);
    const unprotected = await worksheets.unprotectWorkbookStructure(
      next,
      "secret",
    );
    expect(unprotected.revision).toBe(2);
    expect(unprotected.events.at(-1).action).toBe(
      "Unprotect workbook structure",
    );
  });
});
