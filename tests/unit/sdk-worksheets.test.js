import { describe, expect, test } from "vitest";
import {
  addWorksheet,
  renameWorksheet,
  moveWorksheet,
  duplicateWorksheet,
  deleteWorksheet,
  setWorksheetVisibility,
  setWorksheetTabColor,
  renameWorkbook,
  updateWorkbookProperties,
  createWorkbookFromTemplate,
  protectWorkbookStructure,
  unprotectWorkbookStructure,
  updateSheetProtection,
  summarizeWorksheets,
  describeSheetDeletionImpacts,
  findSheetOrThrow,
  setFreezePanes,
  setFreezePanesAtCell,
  setSplitPanes,
  mergeWorksheetCells,
  setHiddenAxisRange,
  readHiddenAxisIntervals,
  applySheetSizeEdits,
  removeOutlineGroup,
  setOutlineGroupCollapsed,
} from "../../src/sdk/worksheets.js";
import {
  createBlankWorkbook,
  getWorkbookCalculation,
} from "../../src/app/03-workbook-model.js";

/** A workbook with "Sheet1" and a second sheet "Data", ids included. */
function twoSheets() {
  const workbook = addWorksheet(createBlankWorkbook("Budget"), "Data");
  return {
    workbook,
    firstId: workbook.sheets[0].id,
    dataId: workbook.sheets[1].id,
  };
}

/** Reads a calculated value from the first sheet of a workbook. */
function valueOf(workbook, sheetId, address) {
  return getWorkbookCalculation(workbook).calculator.getTypedValue(
    sheetId,
    address,
  );
}

describe("addWorksheet", () => {
  test("appends a sheet, records one event, and leaves the original workbook unchanged", () => {
    const original = createBlankWorkbook("Budget");
    const next = addWorksheet(original, "Data");
    expect(next.sheets.map((sheet) => sheet.name)).toEqual(["Sheet1", "Data"]);
    expect(original.sheets).toHaveLength(1);
    expect(original.revision).toBe(0);
    expect(original.events).toEqual([]);
    expect(next.sheets[1].rowCount).toBe(original.sheets[0].rowCount);
    expect(next.revision).toBe(original.revision + 1);
    expect(next.events).toHaveLength(1);
    expect(next.events[0]).toMatchObject({
      action: "Add worksheet",
      actor: "Preparer",
      changes: [],
    });
  });

  test("throws for a duplicate name ignoring case and for an invalid name", () => {
    const workbook = createBlankWorkbook("Budget");
    expect(() => addWorksheet(workbook, "sheet1")).toThrow(
      "A worksheet with that name already exists.",
    );
    expect(() => addWorksheet(workbook, "a/b")).toThrow(
      "Sheet names must contain 1–31 characters",
    );
    expect(() => addWorksheet(workbook, "x".repeat(32))).toThrow(
      "Sheet names must contain 1–31 characters",
    );
  });
});

describe("renameWorksheet", () => {
  test("renames the sheet and rewrites formulas that used the old name", () => {
    const { workbook, firstId, dataId } = twoSheets();
    workbook.sheets[1].cells.B2 = { raw: 21 };
    workbook.sheets[0].cells.A1 = { raw: "=Data!B2*2" };
    const renamed = renameWorksheet(workbook, dataId, "Summary");
    expect(renamed.sheets[1].name).toBe("Summary");
    expect(renamed.sheets[1].id).toBe(dataId);
    expect(renamed.sheets[0].cells.A1.raw).toBe("='Summary'!B2*2");
    expect(valueOf(renamed, firstId, "A1")).toEqual({
      type: "number",
      value: 42,
    });
  });

  test("throws for a name used by another sheet and returns the same object when the name is unchanged", () => {
    const { workbook, dataId } = twoSheets();
    expect(() => renameWorksheet(workbook, dataId, "SHEET1")).toThrow(
      "A worksheet with that name already exists.",
    );
    expect(renameWorksheet(workbook, dataId, "Data")).toBe(workbook);
  });
});

describe("moveWorksheet", () => {
  test("moves a sheet to the given position", () => {
    const { workbook, firstId } = twoSheets();
    const moved = moveWorksheet(workbook, firstId, 1);
    expect(moved.sheets.map((sheet) => sheet.name)).toEqual(["Data", "Sheet1"]);
    expect(workbook.sheets[0].name).toBe("Sheet1");
  });

  test("throws for a position outside the workbook and returns the same object for the current position", () => {
    const { workbook, firstId } = twoSheets();
    expect(() => moveWorksheet(workbook, firstId, 2)).toThrow(
      "Choose a worksheet position inside the workbook.",
    );
    expect(() => moveWorksheet(workbook, firstId, 1.5)).toThrow(
      "Choose a worksheet position inside the workbook.",
    );
    expect(moveWorksheet(workbook, firstId, 0)).toBe(workbook);
  });
});

describe("duplicateWorksheet", () => {
  test("adds a copy after the original with a new id and a numbered name", () => {
    const { workbook, dataId } = twoSheets();
    workbook.sheets[1].cells.A1 = { raw: 7 };
    const copied = duplicateWorksheet(workbook, dataId);
    expect(copied.sheets.map((sheet) => sheet.name)).toEqual([
      "Sheet1",
      "Data",
      "Data (2)",
    ]);
    expect(copied.sheets[2].id).not.toBe(dataId);
    expect(copied.sheets[2].cells.A1).toEqual({ raw: 7 });
    expect(workbook.sheets).toHaveLength(2);
  });

  test("rewrites references to the original sheet so the copy refers to itself", () => {
    const { workbook, dataId } = twoSheets();
    workbook.sheets[1].cells.A1 = { raw: 2 };
    workbook.sheets[1].cells.B1 = { raw: "=Data!A1*3" };
    const copied = duplicateWorksheet(workbook, dataId);
    expect(copied.sheets[2].cells.B1.raw).toBe("='Data (2)'!A1*3");
    expect(valueOf(copied, copied.sheets[2].id, "B1")).toEqual({
      type: "number",
      value: 6,
    });
  });

  test("throws when a formula uses a three-dimensional reference", () => {
    const { workbook, dataId } = twoSheets();
    const threeWay = addWorksheet(workbook, "Other");
    threeWay.sheets[0].cells.A1 = { raw: "=SUM(Data:Other!A1)" };
    expect(() => duplicateWorksheet(threeWay, dataId)).toThrow(
      "Duplicating worksheets with three-dimensional references requires explicit references first.",
    );
  });
});

describe("deleteWorksheet", () => {
  test("throws while the sheet has dependencies, and removes it with confirmation", () => {
    const { workbook, firstId, dataId } = twoSheets();
    workbook.sheets[0].cells.A1 = { raw: "=Data!B2*2" };
    expect(() => deleteWorksheet(workbook, dataId)).toThrow(
      "This worksheet has dependencies. Review them and explicitly confirm deletion.",
    );
    const remaining = deleteWorksheet(workbook, dataId, true);
    expect(remaining.sheets.map((sheet) => sheet.id)).toEqual([firstId]);
    expect(remaining.sheets[0].cells.A1.raw).toContain("#REF!");
  });

  test("throws for the only worksheet and when no visible worksheet would remain", () => {
    const only = createBlankWorkbook("Solo");
    expect(() => deleteWorksheet(only, only.sheets[0].id, true)).toThrow(
      "The final worksheet cannot be deleted.",
    );
    const { workbook, firstId, dataId } = twoSheets();
    const hiddenFirst = setWorksheetVisibility(workbook, firstId, "hidden");
    expect(() => deleteWorksheet(hiddenFirst, dataId, true)).toThrow(
      "At least one worksheet must remain visible.",
    );
  });
});

describe("setWorksheetVisibility", () => {
  test("sets hidden and very hidden, and makes a sheet visible again", () => {
    const { workbook, dataId } = twoSheets();
    const hidden = setWorksheetVisibility(workbook, dataId, "hidden");
    expect(hidden.sheets[1]).toMatchObject({ hidden: true, veryHidden: false });
    const veryHidden = setWorksheetVisibility(workbook, dataId, "veryHidden");
    expect(veryHidden.sheets[1]).toMatchObject({
      hidden: true,
      veryHidden: true,
    });
    const visible = setWorksheetVisibility(veryHidden, dataId, "visible");
    expect(visible.sheets[1]).toMatchObject({
      hidden: false,
      veryHidden: false,
    });
  });

  test("throws for an unknown value or when no visible worksheet would remain, and returns the same object when unchanged", () => {
    const { workbook, firstId, dataId } = twoSheets();
    expect(() => setWorksheetVisibility(workbook, dataId, "gone")).toThrow(
      "Choose visible, hidden or very hidden.",
    );
    const hiddenFirst = setWorksheetVisibility(workbook, firstId, "hidden");
    expect(() => setWorksheetVisibility(hiddenFirst, dataId, "hidden")).toThrow(
      "At least one worksheet must remain visible.",
    );
    expect(setWorksheetVisibility(workbook, dataId, "visible")).toBe(workbook);
  });
});

describe("setWorksheetTabColor", () => {
  test("stores the colour in lower case and clears it when given undefined", () => {
    const { workbook, dataId } = twoSheets();
    const colored = setWorksheetTabColor(workbook, dataId, "#FF0000");
    expect(colored.sheets[1].tabColor).toBe("#ff0000");
    const cleared = setWorksheetTabColor(colored, dataId, undefined);
    expect(cleared.sheets[1]).not.toHaveProperty("tabColor");
  });

  test("throws for a colour without a leading # or a name", () => {
    const { workbook, dataId } = twoSheets();
    expect(() => setWorksheetTabColor(workbook, dataId, "00FF00")).toThrow(
      "Tab colors must use a six-digit hexadecimal color.",
    );
    expect(() => setWorksheetTabColor(workbook, dataId, "red")).toThrow(
      "Tab colors must use a six-digit hexadecimal color.",
    );
  });
});

describe("renameWorkbook", () => {
  test("returns a new workbook with the trimmed name and leaves the original as it was", () => {
    const original = createBlankWorkbook("Budget");
    const renamed = renameWorkbook(original, "  Q2 Close  ");
    expect(renamed.name).toBe("Q2 Close");
    expect(original.name).toBe("Budget");
    expect(renamed).not.toBe(original);
  });

  test("throws for an empty name or one longer than 200 characters", () => {
    const workbook = createBlankWorkbook("Budget");
    expect(() => renameWorkbook(workbook, "   ")).toThrow(
      "Workbook name must contain 1–200 characters.",
    );
    expect(() => renameWorkbook(workbook, "x".repeat(201))).toThrow(
      "Workbook name must contain 1–200 characters.",
    );
  });
});

describe("updateWorkbookProperties", () => {
  test("sets engagement and period in the metadata and keeps the name", () => {
    const workbook = createBlankWorkbook("Budget");
    const updated = updateWorkbookProperties(workbook, {
      engagement: "FY26 audit",
      period: "Q1",
    });
    expect(updated.name).toBe("Budget");
    expect(updated.metadata).toMatchObject({
      engagement: "FY26 audit",
      period: "Q1",
    });
    expect(workbook.metadata).not.toHaveProperty("engagement");
  });

  test("throws for an unknown field and for an engagement over 200 characters; an empty engagement removes it", () => {
    const workbook = createBlankWorkbook("Budget");
    expect(() => updateWorkbookProperties(workbook, { color: "red" })).toThrow(
      "Invalid workbook property fields.",
    );
    expect(() =>
      updateWorkbookProperties(workbook, { engagement: "x".repeat(201) }),
    ).toThrow("Engagement must contain at most 200 characters.");
    const withEngagement = updateWorkbookProperties(workbook, {
      engagement: "FY26",
    });
    const cleared = updateWorkbookProperties(withEngagement, {
      engagement: "   ",
    });
    expect(cleared.metadata).not.toHaveProperty("engagement");
  });
});

describe("createWorkbookFromTemplate", () => {
  test("copies the sheets under a new id and records the template it came from", () => {
    const { workbook } = twoSheets();
    const created = createWorkbookFromTemplate(workbook, {
      name: " Q2 close ",
      engagement: "Acme",
      period: "Q2",
    });
    expect(created.id).not.toBe(workbook.id);
    expect(created.name).toBe("Q2 close");
    expect(created.revision).toBe(0);
    expect(created.sheets.map((sheet) => sheet.name)).toEqual([
      "Sheet1",
      "Data",
    ]);
    expect(created.metadata).toMatchObject({
      engagement: "Acme",
      period: "Q2",
      template: { sourceWorkbookId: workbook.id },
    });
    expect(workbook.name).toBe("Budget");
  });

  test("throws when a required detail is missing", () => {
    const workbook = createBlankWorkbook("Budget");
    expect(() =>
      createWorkbookFromTemplate(workbook, { name: "Q2", period: "Q2" }),
    ).toThrow();
  });
});

describe("protectWorkbookStructure", () => {
  test("protects the structure with a salted hash of the password", async () => {
    const { workbook } = twoSheets();
    const protectedBook = await protectWorkbookStructure(workbook, "secret");
    expect(protectedBook.structureProtection).toMatchObject({ enabled: true });
    expect(protectedBook.structureProtection.passwordHash).toMatch(
      /^[a-f\d]{64}$/,
    );
    expect(protectedBook.structureProtection.salt).toEqual(expect.any(String));
    expect(summarizeWorksheets(protectedBook)[0].structureLocked).toBe(true);
  });

  test("blocks adding a worksheet while protected; an empty password protects without a hash; a non-text password throws", async () => {
    const { workbook } = twoSheets();
    const protectedBook = await protectWorkbookStructure(workbook, "");
    expect(protectedBook.structureProtection).toEqual({ enabled: true });
    expect(() => addWorksheet(protectedBook, "New")).toThrow(
      "Workbook structure is protected.",
    );
    await expect(protectWorkbookStructure(workbook, undefined)).rejects.toThrow(
      "Passwords must be text of at most 200 characters.",
    );
  });
});

describe("unprotectWorkbookStructure", () => {
  test("removes the protection with the right password and rejects the wrong one", async () => {
    const { workbook } = twoSheets();
    const protectedBook = await protectWorkbookStructure(workbook, "secret");
    await expect(
      unprotectWorkbookStructure(protectedBook, "wrong"),
    ).rejects.toThrow("The workbook structure password does not match.");
    const unprotected = await unprotectWorkbookStructure(
      protectedBook,
      "secret",
    );
    expect(unprotected).not.toHaveProperty("structureProtection");
    expect(addWorksheet(unprotected, "New").sheets).toHaveLength(3);
  });

  test("returns the same object when the structure is not protected", async () => {
    const { workbook } = twoSheets();
    expect(await unprotectWorkbookStructure(workbook, "")).toBe(workbook);
  });
});

describe("updateSheetProtection", () => {
  test("protects a sheet with a password and unprotects it with the same password", async () => {
    const { workbook, dataId } = twoSheets();
    const protectedBook = await updateSheetProtection(
      workbook,
      dataId,
      "pw",
      true,
    );
    const state = protectedBook.sheets[1].protection;
    expect(state).toMatchObject({ enabled: true });
    expect(state.salt).toEqual(expect.any(String));
    expect(state.hash).toMatch(/^[a-f\d]{64}$/);
    const unprotected = await updateSheetProtection(
      protectedBook,
      dataId,
      "pw",
      false,
    );
    expect(unprotected.sheets[1].protection).toEqual({ enabled: false });
  });

  test("throws when protecting a protected sheet, unprotecting with the wrong password, or given a non-text password", async () => {
    const { workbook, dataId } = twoSheets();
    const protectedBook = await updateSheetProtection(
      workbook,
      dataId,
      "pw",
      true,
    );
    await expect(
      updateSheetProtection(protectedBook, dataId, "pw", true),
    ).rejects.toThrow(
      "Unprotect this sheet before changing its protection settings.",
    );
    await expect(
      updateSheetProtection(protectedBook, dataId, "bad", false),
    ).rejects.toThrow("The password does not match.");
    await expect(
      updateSheetProtection(workbook, dataId, undefined, true),
    ).rejects.toThrow();
  });
});

describe("summarizeWorksheets", () => {
  test("returns one summary per sheet in tab order with visibility, tab colour and counts", () => {
    const { workbook, dataId } = twoSheets();
    workbook.sheets[0].cells.A1 = { raw: "=1+1" };
    workbook.sheets[0].cells.A2 = { raw: 5 };
    const colored = setWorksheetTabColor(workbook, dataId, "#00ff00");
    expect(summarizeWorksheets(colored)).toEqual([
      expect.objectContaining({
        name: "Sheet1",
        visibility: "visible",
        tabColor: null,
        storedCellCount: 2,
        formulaCellCount: 1,
        structureLocked: false,
      }),
      expect.objectContaining({
        name: "Data",
        visibility: "visible",
        tabColor: "#00ff00",
        storedCellCount: 0,
      }),
    ]);
  });

  test("lists only the sheets in allowedSheetIds", () => {
    const { workbook, dataId } = twoSheets();
    const only = summarizeWorksheets(workbook, { allowedSheetIds: [dataId] });
    expect(only.map((summary) => summary.name)).toEqual(["Data"]);
  });
});

describe("describeSheetDeletionImpacts", () => {
  test("returns no impacts when no other sheet refers to the sheet", () => {
    const { workbook, dataId } = twoSheets();
    expect(describeSheetDeletionImpacts(workbook, dataId)).toEqual([]);
  });

  test("describes each formula that would become #REF!", () => {
    const { workbook, dataId } = twoSheets();
    workbook.sheets[0].cells.A1 = { raw: "=Data!B2*2" };
    expect(describeSheetDeletionImpacts(workbook, dataId)).toEqual([
      {
        kind: "formula",
        label: "Sheet1!A1 will contain #REF!",
        sheetId: workbook.sheets[0].id,
        address: "A1",
      },
    ]);
  });
});

describe("findSheetOrThrow", () => {
  test("returns the sheet with the given id", () => {
    const { workbook, dataId } = twoSheets();
    expect(findSheetOrThrow(workbook, dataId).name).toBe("Data");
  });

  test("throws when no sheet has the id", () => {
    const { workbook } = twoSheets();
    expect(() => findSheetOrThrow(workbook, "missing")).toThrow(
      "Sheet not found: missing",
    );
  });
});

describe("setFreezePanes", () => {
  test("freezes the given numbers of rows and columns", () => {
    const { workbook, dataId } = twoSheets();
    const frozen = setFreezePanes(workbook, dataId, 2, 1);
    expect(frozen.sheets[1]).toMatchObject({ frozenRows: 2, frozenCols: 1 });
  });

  test("throws for a negative count and removes split panes when freezing", () => {
    const { workbook, dataId } = twoSheets();
    expect(() => setFreezePanes(workbook, dataId, -1, 0)).toThrow(
      "Frozen boundaries must be valid leading row and column counts.",
    );
    const split = setSplitPanes(workbook, dataId, {
      horizontal: 50,
      row: 1,
      col: 0,
    });
    const frozen = setFreezePanes(split, dataId, 1, 0);
    expect(frozen.sheets[1]).not.toHaveProperty("splitPanes");
    expect(frozen.sheets[1].frozenRows).toBe(1);
  });
});

describe("setFreezePanesAtCell", () => {
  test("freezes the rows above and the columns left of the cell", () => {
    const { workbook, dataId } = twoSheets();
    const frozen = setFreezePanesAtCell(workbook, dataId, "B3");
    expect(frozen.sheets[1]).toMatchObject({ frozenRows: 2, frozenCols: 1 });
    const top = setFreezePanesAtCell(workbook, dataId, "A1");
    expect(top.sheets[1]).toMatchObject({ frozenRows: 0, frozenCols: 0 });
  });

  test("throws for an invalid cell address", () => {
    const { workbook, dataId } = twoSheets();
    expect(() => setFreezePanesAtCell(workbook, dataId, "??")).toThrow(
      "Invalid cell address: ??",
    );
  });
});

describe("setSplitPanes", () => {
  test("sets the split and clears frozen panes", () => {
    const { workbook, dataId } = twoSheets();
    const frozen = setFreezePanes(workbook, dataId, 1, 1);
    const split = setSplitPanes(frozen, dataId, {
      horizontal: 50,
      row: 1,
      col: 0,
    });
    expect(split.sheets[1]).toMatchObject({
      splitPanes: { horizontal: 50, row: 1, col: 0 },
      frozenRows: 0,
      frozenCols: 0,
    });
  });

  test("removes the split when given undefined and throws for a divider outside 20 to 80 or no divider", () => {
    const { workbook, dataId } = twoSheets();
    const split = setSplitPanes(workbook, dataId, {
      horizontal: 50,
      row: 1,
      col: 0,
    });
    expect(
      setSplitPanes(split, dataId, undefined).sheets[1],
    ).not.toHaveProperty("splitPanes");
    expect(() =>
      setSplitPanes(workbook, dataId, { horizontal: 10, row: 1, col: 0 }),
    ).toThrow("Split dividers must be between 20% and 80%.");
    expect(() => setSplitPanes(workbook, dataId, { row: 1, col: 0 })).toThrow(
      "Choose a horizontal split, vertical split, or both.",
    );
  });
});

describe("mergeWorksheetCells", () => {
  test("merges a range, records it and centres the anchor cell", () => {
    const { workbook, dataId } = twoSheets();
    const merged = mergeWorksheetCells(workbook, dataId, "B2:C3");
    expect(merged.sheets[1].merges).toEqual(["B2:C3"]);
    expect(merged.sheets[1].cells.B2.style).toEqual({ align: "center" });
  });

  test("throws when the range overlaps a merge, and unmerging removes it", () => {
    const { workbook, dataId } = twoSheets();
    const merged = mergeWorksheetCells(workbook, dataId, "B2:C3");
    expect(() => mergeWorksheetCells(merged, dataId, "C3:D4")).toThrow(
      "Unmerge overlapping areas first.",
    );
    const unmerged = mergeWorksheetCells(merged, dataId, "B2:C3", true);
    expect(unmerged.sheets[1].merges).toEqual([]);
  });
});

describe("setHiddenAxisRange", () => {
  test("hides rows from start to end inclusive", () => {
    const { workbook, dataId } = twoSheets();
    const hidden = setHiddenAxisRange(workbook, dataId, "row", 1, 3, true);
    expect(readHiddenAxisIntervals(hidden.sheets[1], "row")).toEqual([
      { start: 1, end: 3 },
    ]);
  });

  test("throws for an unknown axis, a non-boolean state or a range outside the grid", () => {
    const { workbook, dataId } = twoSheets();
    expect(() =>
      setHiddenAxisRange(workbook, dataId, "cell", 0, 1, true),
    ).toThrow("Choose rows or columns.");
    expect(() =>
      setHiddenAxisRange(workbook, dataId, "row", 0, 1, "yes"),
    ).toThrow("Hidden state must be boolean.");
    expect(() =>
      setHiddenAxisRange(workbook, dataId, "column", 0, 99999, true),
    ).toThrow(
      "Hidden intervals require valid zero-based inclusive axis bounds.",
    );
  });
});

describe("readHiddenAxisIntervals", () => {
  test("merges adjacent hidden runs into one interval", () => {
    const { workbook, dataId } = twoSheets();
    let next = setHiddenAxisRange(workbook, dataId, "row", 0, 2, true);
    next = setHiddenAxisRange(next, dataId, "row", 3, 5, true);
    expect(readHiddenAxisIntervals(next.sheets[1], "row")).toEqual([
      { start: 0, end: 5 },
    ]);
  });

  test("returns an empty list when nothing is hidden and throws for an unknown axis", () => {
    const { workbook } = twoSheets();
    expect(readHiddenAxisIntervals(workbook.sheets[1], "row")).toEqual([]);
    expect(() => readHiddenAxisIntervals(workbook.sheets[1], "cell")).toThrow(
      "Choose rows or columns.",
    );
  });
});

describe("applySheetSizeEdits", () => {
  test("sets the default row height and column width in pixels", () => {
    const { workbook, dataId } = twoSheets();
    const sized = applySheetSizeEdits(workbook, dataId, {
      rowHeight: 30,
      columnWidth: 120,
    });
    expect(sized.sheets[1]).toMatchObject({
      defaultRowHeight: 30,
      defaultColumnWidth: 120,
    });
  });

  test("returns the same object when nothing changes and throws for an out-of-range height, an unknown key, a stale snapshot or a read-only workbook", () => {
    const { workbook, dataId } = twoSheets();
    expect(applySheetSizeEdits(workbook, dataId, {})).toBe(workbook);
    expect(() =>
      applySheetSizeEdits(workbook, dataId, { rowHeight: 20 }),
    ).toThrow("Default row height must be 21–546 pixels.");
    expect(() => applySheetSizeEdits(workbook, dataId, { color: 1 })).toThrow(
      "Choose valid worksheet size settings.",
    );
    expect(() =>
      applySheetSizeEdits(
        workbook,
        dataId,
        { rowHeight: 30 },
        {
          expectedSnapshot: {
            workbookId: workbook.id,
            revision: 99,
            snapshotGuard: "stale",
          },
        },
      ),
    ).toThrow("The workbook changed after these sizes were opened.");
    expect(() =>
      applySheetSizeEdits(
        workbook,
        dataId,
        { rowHeight: 30 },
        { readOnly: true },
      ),
    ).toThrow("This workbook is read-only.");
  });
});

describe("applySheetSizeEdits resets", () => {
  test("clears custom row heights and column widths when asked, and keeps the others", () => {
    const { workbook, dataId } = twoSheets();
    const withOverrides = {
      ...workbook,
      sheets: workbook.sheets.map((sheet) =>
        sheet.id === dataId
          ? { ...sheet, rowHeights: { 2: 40 }, columnWidths: { 1: 150 } }
          : sheet,
      ),
    };
    const rowsReset = applySheetSizeEdits(withOverrides, dataId, {
      resetRowOverrides: true,
    });
    expect(rowsReset.sheets[1]).not.toHaveProperty("rowHeights");
    expect(rowsReset.sheets[1].columnWidths).toEqual({ 1: 150 });
    const bothReset = applySheetSizeEdits(withOverrides, dataId, {
      resetRowOverrides: true,
      resetColumnOverrides: true,
    });
    expect(bothReset.sheets[1]).not.toHaveProperty("columnWidths");
    expect(withOverrides.sheets[1].rowHeights).toEqual({ 2: 40 });
  });
});

/** A two-sheet workbook whose Data sheet has one outline group, "g1", on rows 1 to 3. */
function withOutlineGroup() {
  const { workbook, dataId } = twoSheets();
  const grouped = {
    ...workbook,
    sheets: workbook.sheets.map((sheet) =>
      sheet.id === dataId
        ? {
            ...sheet,
            outlineGroups: [
              { id: "g1", axis: "row", start: 1, end: 3, collapsed: false },
            ],
          }
        : sheet,
    ),
  };
  return { workbook: grouped, dataId };
}

describe("removeOutlineGroup", () => {
  test("removes the group with the given id", () => {
    const { workbook, dataId } = withOutlineGroup();
    const removed = removeOutlineGroup(workbook, dataId, "g1");
    expect(removed.sheets[1].outlineGroups).toEqual([]);
    expect(workbook.sheets[1].outlineGroups).toHaveLength(1);
  });

  test("throws when no group has the id", () => {
    const { workbook, dataId } = withOutlineGroup();
    expect(() => removeOutlineGroup(workbook, dataId, "nope")).toThrow(
      "Outline group not found.",
    );
  });
});

describe("setOutlineGroupCollapsed", () => {
  test("collapses and expands the group", () => {
    const { workbook, dataId } = withOutlineGroup();
    const collapsed = setOutlineGroupCollapsed(workbook, dataId, "g1", true);
    expect(collapsed.sheets[1].outlineGroups[0].collapsed).toBe(true);
    const expanded = setOutlineGroupCollapsed(collapsed, dataId, "g1", false);
    expect(expanded.sheets[1].outlineGroups[0].collapsed).toBe(false);
  });

  test("throws for a non-boolean state and for a missing group", () => {
    const { workbook, dataId } = withOutlineGroup();
    expect(() =>
      setOutlineGroupCollapsed(workbook, dataId, "g1", "yes"),
    ).toThrow("Collapsed state must be boolean.");
    expect(() =>
      setOutlineGroupCollapsed(workbook, dataId, "nope", true),
    ).toThrow("Outline group not found.");
  });
});
