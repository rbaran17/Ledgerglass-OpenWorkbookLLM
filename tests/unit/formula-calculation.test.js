import { describe, expect, test } from "vitest";
import {
  createBlankWorkbook,
  getWorkbookCalculation,
} from "../../src/app/03-workbook-model.js";

/**
 * Builds a workbook whose first sheet holds `cells` (address -> raw input) and,
 * optionally, a second sheet named "Other" holding `otherCells`.
 */
function calculate(cells, otherCells = null) {
  const workbook = createBlankWorkbook("Unit");
  const sheet = workbook.sheets[0];
  for (const [address, raw] of Object.entries(cells))
    sheet.cells[address] = { raw };
  let other = null;
  if (otherCells) {
    other = {
      id: "sheet-other",
      name: "Other",
      rowCount: 1048576,
      colCount: 16384,
      cells: {},
      frozenRows: 0,
    };
    for (const [address, raw] of Object.entries(otherCells))
      other.cells[address] = { raw };
    workbook.sheets.push(other);
  }
  const { calculator, report } = getWorkbookCalculation(workbook);
  return {
    report,
    value: (address) => calculator.getTypedValue(sheet.id, address),
  };
}

describe("arithmetic and operator precedence", () => {
  test("multiplication binds tighter than addition", () => {
    const { value } = calculate({ A1: "=2+3*4" });
    expect(value("A1")).toMatchObject({ type: "number", value: 14 });
  });

  test("parentheses override the default precedence", () => {
    const { value } = calculate({ A1: "=(2+3)*4" });
    expect(value("A1")).toMatchObject({ type: "number", value: 20 });
  });

  test("subtraction is left-associative", () => {
    const { value } = calculate({ A1: "=10-4-3" });
    expect(value("A1")).toMatchObject({ type: "number", value: 3 });
  });

  test("exponentiation is left-associative, as in Excel", () => {
    const { value } = calculate({ A1: "=2^3^2" });
    expect(value("A1")).toMatchObject({ type: "number", value: 64 });
  });

  test("unary minus binds tighter than exponentiation, as in Excel", () => {
    const { value } = calculate({ A1: "=-2^2" });
    expect(value("A1")).toMatchObject({ type: "number", value: 4 });
  });
});

describe("cell and range references", () => {
  test("a formula reads the value of a referenced cell", () => {
    const { value } = calculate({ A1: "=2", A2: "=A1*21" });
    expect(value("A2")).toMatchObject({ type: "number", value: 42 });
  });

  test("SUM over a range adds every numeric cell in it", () => {
    const { value } = calculate({ A1: "=2", A2: "=3", A3: "=SUM(A1:A2)" });
    expect(value("A3")).toMatchObject({ type: "number", value: 5 });
  });

  test("a chain of dependent cells recalculates in order", () => {
    const { value, report } = calculate({
      A1: "=2",
      A2: "=A1*3",
      A3: "=A2+A1",
      A4: "=A3*A3",
    });
    expect(value("A2")).toMatchObject({ value: 6 });
    expect(value("A3")).toMatchObject({ value: 8 });
    expect(value("A4")).toMatchObject({ value: 64 });
    expect(report.cycles).toHaveLength(0);
  });

  test("a formula can reference a cell on another sheet", () => {
    const { value } = calculate(
      { A1: "=Other!B2*2" },
      { B1: "=10", B2: "=B1+5" },
    );
    expect(value("A1")).toMatchObject({ type: "number", value: 30 });
  });

  test("a reference to a missing sheet returns #REF!", () => {
    const { value } = calculate({ A1: "=Missing!A1" });
    expect(value("A1")).toMatchObject({ type: "error", value: "#REF!" });
  });
});

describe("text, comparison and concatenation", () => {
  test("ampersand joins text and numbers", () => {
    const { value } = calculate({ A1: '=2&"x"', A2: '="ab"&"c"' });
    expect(value("A1")).toMatchObject({ type: "text", value: "2x" });
    expect(value("A2")).toMatchObject({ type: "text", value: "abc" });
  });

  test("ampersand binds looser than plus", () => {
    const { value } = calculate({ A1: '="a"&1+2' });
    expect(value("A1")).toMatchObject({ type: "text", value: "a3" });
  });

  test("comparison binds looser than concatenation", () => {
    const { value } = calculate({ A1: '="a"&"b"="ab"' });
    expect(value("A1")).toMatchObject({ type: "boolean", value: true });
  });

  test("numeric comparison operators return booleans", () => {
    const { value } = calculate({
      A1: "=10>9",
      A2: "=1<>1",
      A3: "=1+2=3",
    });
    expect(value("A1")).toMatchObject({ type: "boolean", value: true });
    expect(value("A2")).toMatchObject({ type: "boolean", value: false });
    expect(value("A3")).toMatchObject({ type: "boolean", value: true });
  });

  test("text comparison ignores letter case", () => {
    const { value } = calculate({ A1: '="a"="A"' });
    expect(value("A1")).toMatchObject({ type: "boolean", value: true });
  });

  test("text compares as greater than any number", () => {
    const { value } = calculate({ A1: '="10">9' });
    expect(value("A1")).toMatchObject({ type: "boolean", value: true });
  });
});

describe("logical functions", () => {
  test("IF returns the branch selected by its condition", () => {
    const { value } = calculate({ A1: '=IF(1>2,"yes","no")' });
    expect(value("A1")).toMatchObject({ type: "text", value: "no" });
  });

  test("IF does not evaluate the branch it does not take", () => {
    const { value } = calculate({ A1: "=IF(TRUE,1,1/0)" });
    expect(value("A1")).toMatchObject({ type: "number", value: 1 });
  });

  test("AND and OR combine booleans", () => {
    const { value } = calculate({
      A1: "=AND(TRUE,FALSE)",
      A2: "=OR(FALSE,TRUE)",
    });
    expect(value("A1")).toMatchObject({ type: "boolean", value: false });
    expect(value("A2")).toMatchObject({ type: "boolean", value: true });
  });

  test("NOT inverts a boolean and treats non-zero as TRUE", () => {
    const { value } = calculate({ A1: "=NOT(FALSE)", A2: "=NOT(1)" });
    expect(value("A1")).toMatchObject({ type: "boolean", value: true });
    expect(value("A2")).toMatchObject({ type: "boolean", value: false });
  });
});

describe("error values and propagation", () => {
  test("division by zero returns #DIV/0!", () => {
    const { value } = calculate({ A1: "=1/0" });
    expect(value("A1")).toMatchObject({ type: "error", value: "#DIV/0!" });
  });

  test("an unknown function name returns #NAME?", () => {
    const { value } = calculate({ A1: "=UNKNOWNFN(1)" });
    expect(value("A1")).toMatchObject({ type: "error", value: "#NAME?" });
  });

  test("arithmetic on text returns #VALUE!", () => {
    const { value } = calculate({ A1: '="x"+1' });
    expect(value("A1")).toMatchObject({ type: "error", value: "#VALUE!" });
  });

  test("an error in a referenced cell propagates through arithmetic and SUM", () => {
    const { value } = calculate({
      A1: "=1/0",
      A2: "=A1+1",
      A3: "=SUM(A1:A2)",
    });
    expect(value("A2")).toMatchObject({ type: "error", value: "#DIV/0!" });
    expect(value("A3")).toMatchObject({ type: "error", value: "#DIV/0!" });
  });

  test("an error in an IF condition propagates", () => {
    const { value } = calculate({ A1: "=IF(1/0,1,2)" });
    expect(value("A1")).toMatchObject({ type: "error", value: "#DIV/0!" });
  });

  test("IFERROR returns the fallback when the first argument is an error", () => {
    const { value } = calculate({
      A1: '=IFERROR(1/0,"fallback")',
      A2: "=IFERROR(2+2,0)",
    });
    expect(value("A1")).toMatchObject({ type: "text", value: "fallback" });
    expect(value("A2")).toMatchObject({ type: "number", value: 4 });
  });

  test("IFERROR catches an error inherited from another cell", () => {
    const { value } = calculate({ A1: '="x"+1', A2: "=IFERROR(A1,0)" });
    expect(value("A2")).toMatchObject({ type: "number", value: 0 });
  });
});

describe("empty cells", () => {
  test("an empty cell counts as zero in arithmetic", () => {
    const { value } = calculate({ A1: "=A9+1", A2: "=5-A9" });
    expect(value("A1")).toMatchObject({ type: "number", value: 1 });
    expect(value("A2")).toMatchObject({ type: "number", value: 5 });
  });
});

describe("circular references", () => {
  test("a circular pair is reported as #CYCLE! instead of hanging", () => {
    const { value, report } = calculate({ A1: "=A2+1", A2: "=A1+1" });
    expect(value("A1")).toMatchObject({ type: "error", value: "#CYCLE!" });
    expect(value("A2")).toMatchObject({ type: "error", value: "#CYCLE!" });
    expect(report.cycles).toHaveLength(1);
  });
});
