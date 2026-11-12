import { describe, expect, test } from "vitest";
import {
  applyCellPatchTransaction,
  buildCanonicalWorkbookSnapshot,
  createBlankWorkbook,
  formatCellValueWithNumberFormat,
} from "../../src/app/03-workbook-model.js";

const SHEET_ID = (workbook) => workbook.sheets[0].id;

function patch(workbook, address, raw) {
  return { sheetId: SHEET_ID(workbook), address, cell: { raw } };
}

function format(value, pattern) {
  return formatCellValueWithNumberFormat(value, {
    category: "custom",
    locale: "en-US",
    pattern,
  }).text;
}

describe("cell patch transactions", () => {
  test("applying a patch sets the cell in the returned workbook", () => {
    const workbook = createBlankWorkbook("Edits");
    const next = applyCellPatchTransaction(workbook, [
      patch(workbook, "A1", "42"),
    ]);
    expect(next.sheets[0].cells.A1).toMatchObject({ raw: "42" });
  });

  test("the original workbook is not mutated", () => {
    const workbook = createBlankWorkbook("Edits");
    const before = JSON.stringify(workbook);
    const next = applyCellPatchTransaction(workbook, [
      patch(workbook, "B2", "hello"),
    ]);
    expect(next).not.toBe(workbook);
    expect(JSON.stringify(workbook)).toBe(before);
    expect(workbook.sheets[0].cells).toEqual({});
  });

  test("the revision advances by one", () => {
    const workbook = createBlankWorkbook("Edits");
    const next = applyCellPatchTransaction(workbook, [
      patch(workbook, "A1", "1"),
    ]);
    expect(next.revision).toBe(workbook.revision + 1);
  });

  test("a transaction records one event with actor, action and changes", () => {
    const workbook = createBlankWorkbook("Edits");
    const next = applyCellPatchTransaction(
      workbook,
      [patch(workbook, "A1", "7"), patch(workbook, "A2", "8")],
      "Alice",
      "Enter figures",
    );
    expect(next.events).toHaveLength(1);
    const [event] = next.events;
    expect(event).toMatchObject({
      actor: "Alice",
      action: "Enter figures",
    });
    expect(typeof event.id).toBe("string");
    expect(typeof event.timestamp).toBe("string");
    expect(event.changes.map((change) => change.address)).toEqual(["A1", "A2"]);
    expect(event.changes[0].after).toMatchObject({ raw: "7" });
    expect(event.changes[0].before).toBeUndefined();
  });

  test("successive transactions append events in order", () => {
    const first = createBlankWorkbook("Edits");
    const second = applyCellPatchTransaction(first, [patch(first, "A1", "1")]);
    const third = applyCellPatchTransaction(
      second,
      [patch(second, "A1", "2")],
      "Bob",
      "Change A1",
    );
    expect(third.events.map((event) => event.actor)).toEqual([
      "Preparer",
      "Bob",
    ]);
    expect(third.events[1].changes[0].before).toMatchObject({ raw: "1" });
    expect(third.sheets[0].cells.A1).toMatchObject({ raw: "2" });
  });

  test("an empty patch list returns the same workbook", () => {
    const workbook = createBlankWorkbook("Edits");
    expect(applyCellPatchTransaction(workbook, [])).toBe(workbook);
  });

  test("more than 500 cell changes is rejected", () => {
    const workbook = createBlankWorkbook("Edits");
    const patches = Array.from({ length: 501 }, (_, index) =>
      patch(workbook, `A${index + 1}`, "1"),
    );
    expect(() => applyCellPatchTransaction(workbook, patches)).toThrow(
      /at most 500 cell changes/,
    );
  });

  test("a patch for an unknown sheet is rejected", () => {
    const workbook = createBlankWorkbook("Edits");
    expect(() =>
      applyCellPatchTransaction(workbook, [
        { sheetId: "no-such-sheet", address: "A1", cell: { raw: "1" } },
      ]),
    ).toThrow(/Sheet not found/);
  });

  test("the same cell cannot be changed twice in one transaction", () => {
    const workbook = createBlankWorkbook("Edits");
    expect(() =>
      applyCellPatchTransaction(workbook, [
        patch(workbook, "C3", "1"),
        patch(workbook, "C3", "2"),
      ]),
    ).toThrow(/cannot change the same cell twice/);
  });

  test("equal workbooks produce equal canonical snapshots", () => {
    const workbook = createBlankWorkbook("Snapshot");
    const edited = applyCellPatchTransaction(workbook, [
      patch(workbook, "A1", "3"),
    ]);
    const copy = JSON.parse(JSON.stringify(edited));
    expect(JSON.stringify(buildCanonicalWorkbookSnapshot(copy))).toBe(
      JSON.stringify(buildCanonicalWorkbookSnapshot(edited)),
    );
  });
});

describe("number formatting with Excel format codes", () => {
  test("0.00 always shows two decimals", () => {
    expect(format(1234.5, "0.00")).toBe("1234.50");
    expect(format(1, "0.00")).toBe("1.00");
  });

  test("#,##0 groups thousands and drops decimals", () => {
    expect(format(1234567.891, "#,##0")).toBe("1,234,568");
    expect(format(0, "#,##0")).toBe("0");
  });

  test("#,##0.00 groups thousands with two decimals", () => {
    expect(format(1234.5, "#,##0.00")).toBe("1,234.50");
  });

  test("0% multiplies by 100 and rounds to a whole percent", () => {
    expect(format(0.256, "0%")).toBe("26%");
  });

  test("0.0% keeps one decimal place", () => {
    expect(format(0.256, "0.0%")).toBe("25.6%");
  });

  test("negative numbers take a leading minus sign", () => {
    expect(format(-1234.5, "#,##0.00")).toBe("-1,234.50");
    expect(format(-0.5, "0.00")).toBe("-0.50");
  });

  test("a second section can show negatives in parentheses", () => {
    expect(format(-1234.5, "#,##0.00;(#,##0.00)")).toBe("(1,234.50)");
    expect(format(1234.5, "#,##0.00;(#,##0.00)")).toBe("1,234.50");
  });

  test("yyyy-mm-dd renders a fixed serial as an ISO date", () => {
    expect(format(45000, "yyyy-mm-dd")).toBe("2023-03-15");
  });
});
