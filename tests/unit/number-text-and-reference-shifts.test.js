import { describe, expect, test } from "vitest";
import {
  createBlankWorkbook,
  getWorkbookCalculation,
  shiftFormulaReferences,
} from "../../src/app/03-workbook-model.js";

function valueOf(formula) {
  const workbook = createBlankWorkbook("Unit");
  const sheet = workbook.sheets[0];
  sheet.cells.A1 = { raw: formula };
  return getWorkbookCalculation(workbook).calculator.getTypedValue(
    sheet.id,
    "A1",
  ).value;
}

describe("numbers converted to text use 15 significant digits", () => {
  test("the & operator hides binary rounding noise", () => {
    expect(valueOf('=0.1+0.2&""')).toBe("0.3");
  });

  test("the & operator keeps 15 digits of a repeating fraction", () => {
    expect(valueOf('=1/3&""')).toBe("0.333333333333333");
    expect(valueOf('=2/3&""')).toBe("0.666666666666667");
    expect(valueOf('=-1/3&""')).toBe("-0.333333333333333");
  });

  test("the & operator leaves short numbers, text and booleans alone", () => {
    expect(valueOf('=1234567.891&""')).toBe("1234567.891");
    expect(valueOf('="a"&1&TRUE')).toBe("a1TRUE");
  });

  test("CONCAT and LEN agree with the & operator", () => {
    expect(valueOf("=CONCAT(0.1+0.2)")).toBe("0.3");
    expect(valueOf("=LEN(1/3)")).toBe(17);
  });
});

describe("shifting a range that leaves the grid", () => {
  test("a cell range with one corner off the grid becomes #REF!", () => {
    expect(shiftFormulaReferences("=B1:C2", 0, -2)).toBe("=#REF!");
    expect(shiftFormulaReferences("=A1:XFD1", 0, 1)).toBe("=#REF!");
  });

  test("row and column ranges pushed off the grid become #REF!", () => {
    expect(shiftFormulaReferences("=1:3", -1, 0)).toBe("=#REF!");
    expect(shiftFormulaReferences("=A:B", 0, -1)).toBe("=#REF!");
  });

  test("a range that stays on the grid still shifts both corners", () => {
    expect(shiftFormulaReferences("=B1:C2", 1, 1)).toBe("=C2:D3");
    expect(shiftFormulaReferences("=SUM(B2:C3)+D4", -1, -1)).toBe(
      "=SUM(A1:B2)+C3",
    );
  });

  test("absolute corners stay put while relative corners move", () => {
    expect(shiftFormulaReferences("=$B$1:C2", 5, 5)).toBe("=$B$1:H7");
  });
});
