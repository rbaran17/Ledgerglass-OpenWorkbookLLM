import { describe, expect, test } from "vitest";
import {
  createBlankWorkbook,
  getWorkbookCalculation,
} from "../../src/app/03-workbook-model.js";

// Builds a blank workbook from { address: raw } pairs and returns a reader for
// the calculated typed values. Numbers are passed as JS numbers, text as strings
// that start with a letter, and formulas as strings beginning with "=".
function calculate(cells) {
  const workbook = createBlankWorkbook("Unit");
  const sheet = workbook.sheets[0];
  for (const [address, raw] of Object.entries(cells))
    sheet.cells[address] = { raw };
  const { calculator } = getWorkbookCalculation(workbook);
  return {
    value: (address) => calculator.getTypedValue(sheet.id, address),
  };
}

describe("aggregate functions", () => {
  test("SUM adds a range of numbers", () => {
    const { value } = calculate({ A1: 1, A2: 2, A3: 3, B1: "=SUM(A1:A3)" });
    expect(value("B1")).toMatchObject({ type: "number", value: 6 });
  });

  test("SUM ignores text cells inside a range", () => {
    const { value } = calculate({
      A1: 10,
      A2: 20,
      A3: "apple",
      A4: 30,
      B1: "=SUM(A1:A4)",
    });
    expect(value("B1")).toMatchObject({ type: "number", value: 60 });
  });

  test("AVERAGE divides by the count of numeric cells", () => {
    const { value } = calculate({ A1: 1, A2: 2, A3: 4, B1: "=AVERAGE(A1:A3)" });
    expect(value("B1")).toMatchObject({ type: "number", value: 7 / 3 });
  });

  test("AVERAGE skips text cells in a range", () => {
    const { value } = calculate({
      A1: 10,
      A2: 20,
      A3: "apple",
      A4: 30,
      B1: "=AVERAGE(A1:A4)",
    });
    expect(value("B1")).toMatchObject({ type: "number", value: 20 });
  });

  test("MIN and MAX handle negative numbers", () => {
    const { value } = calculate({
      A1: 5,
      A2: -2,
      A3: 9,
      B1: "=MIN(A1:A3)",
      B2: "=MAX(A1:A3)",
    });
    expect(value("B1")).toMatchObject({ type: "number", value: -2 });
    expect(value("B2")).toMatchObject({ type: "number", value: 9 });
  });

  test("COUNT counts only numeric cells", () => {
    const { value } = calculate({
      A1: 1,
      A2: "text",
      A3: 4,
      A4: "=A1*2",
      B1: "=COUNT(A1:A4)",
    });
    expect(value("B1")).toMatchObject({ type: "number", value: 3 });
  });

  test("COUNTA counts every non-empty cell, including text", () => {
    const { value } = calculate({
      A1: 10,
      A2: 20,
      A3: "apple",
      A4: 30,
      B1: "=COUNTA(A1:A5)",
    });
    expect(value("B1")).toMatchObject({ type: "number", value: 4 });
  });

  test("SUMIF sums values that meet a numeric criterion", () => {
    const { value } = calculate({
      A1: 10,
      A2: 20,
      A3: "apple",
      A4: 30,
      B1: '=SUMIF(A1:A4,">15")',
    });
    expect(value("B1")).toMatchObject({ type: "number", value: 50 });
  });

  test("SUMIF with a separate sum range adds matching rows", () => {
    const { value } = calculate({
      A1: "a",
      B1: 1,
      A2: "b",
      B2: 2,
      A3: "b",
      B3: 3,
      C1: '=SUMIF(A1:A3,"b",B1:B3)',
    });
    expect(value("C1")).toMatchObject({ type: "number", value: 5 });
  });

  test("COUNTIF counts with comparison operators and exact text", () => {
    const { value } = calculate({
      A1: 10,
      A2: 20,
      A3: "apple",
      A4: 30,
      B1: '=COUNTIF(A1:A4,">=20")',
      B2: '=COUNTIF(A1:A4,"apple")',
    });
    expect(value("B1")).toMatchObject({ type: "number", value: 2 });
    expect(value("B2")).toMatchObject({ type: "number", value: 1 });
  });

  test("SUMPRODUCT multiplies paired cells and adds the results", () => {
    const { value } = calculate({
      A1: 1,
      A2: 2,
      B1: 3,
      B2: 4,
      C1: "=SUMPRODUCT(A1:A2,B1:B2)",
    });
    expect(value("C1")).toMatchObject({ type: "number", value: 11 });
  });
});

describe("rounding and arithmetic functions", () => {
  test("ROUND rounds halves away from zero", () => {
    const { value } = calculate({
      A1: "=ROUND(2.5,0)",
      A2: "=ROUND(-2.5,0)",
      A3: "=ROUND(-1.5,0)",
    });
    expect(value("A1")).toMatchObject({ type: "number", value: 3 });
    expect(value("A2")).toMatchObject({ type: "number", value: -3 });
    expect(value("A3")).toMatchObject({ type: "number", value: -2 });
  });

  test("ROUND accepts a negative digit count to round to hundreds", () => {
    const { value } = calculate({
      A1: "=ROUND(1234.5678,2)",
      A2: "=ROUND(1234.5678,-2)",
    });
    expect(value("A1")).toMatchObject({ type: "number", value: 1234.57 });
    expect(value("A2")).toMatchObject({ type: "number", value: 1200 });
  });

  test("ROUNDUP moves away from zero and ROUNDDOWN moves toward zero", () => {
    const { value } = calculate({
      A1: "=ROUNDUP(2.1,0)",
      A2: "=ROUNDUP(-2.1,0)",
      A3: "=ROUNDDOWN(-2.9,0)",
      A4: "=ROUNDDOWN(3.14159,2)",
    });
    expect(value("A1")).toMatchObject({ type: "number", value: 3 });
    expect(value("A2")).toMatchObject({ type: "number", value: -3 });
    expect(value("A3")).toMatchObject({ type: "number", value: -2 });
    expect(value("A4")).toMatchObject({ type: "number", value: 3.14 });
  });

  test("INT rounds down toward negative infinity", () => {
    const { value } = calculate({
      A1: "=INT(7.9)",
      A2: "=INT(-2.5)",
    });
    expect(value("A1")).toMatchObject({ type: "number", value: 7 });
    expect(value("A2")).toMatchObject({ type: "number", value: -3 });
  });

  test("MOD takes the sign of the divisor", () => {
    const { value } = calculate({
      A1: "=MOD(-3,2)",
      A2: "=MOD(10,4)",
      A3: "=MOD(5.5,2)",
    });
    expect(value("A1")).toMatchObject({ type: "number", value: 1 });
    expect(value("A2")).toMatchObject({ type: "number", value: 2 });
    expect(value("A3")).toMatchObject({ type: "number", value: 1.5 });
  });
});

describe("text functions", () => {
  test("LEFT, RIGHT and MID extract parts of a string", () => {
    const { value } = calculate({
      A1: "Hello",
      B1: "=LEFT(A1,2)",
      B2: "=RIGHT(A1,3)",
      B3: "=MID(A1,2,3)",
    });
    expect(value("B1")).toMatchObject({ type: "text", value: "He" });
    expect(value("B2")).toMatchObject({ type: "text", value: "llo" });
    expect(value("B3")).toMatchObject({ type: "text", value: "ell" });
  });

  test("LEN counts characters and LEN of an empty string is zero", () => {
    const { value } = calculate({
      A1: "Hello",
      B1: "=LEN(A1)",
      B2: '=LEN("")',
    });
    expect(value("B1")).toMatchObject({ type: "number", value: 5 });
    expect(value("B2")).toMatchObject({ type: "number", value: 0 });
  });

  test("TRIM removes outer spaces and collapses inner runs of spaces", () => {
    const { value } = calculate({
      A1: '=TRIM("  a   b  ")',
      A2: '=TRIM(" x ")',
    });
    expect(value("A1")).toMatchObject({ type: "text", value: "a b" });
    expect(value("A2")).toMatchObject({ type: "text", value: "x" });
  });

  test("UPPER and LOWER change letter case", () => {
    const { value } = calculate({
      A1: "Hello",
      B1: "=UPPER(A1)",
      B2: "=LOWER(A1)",
    });
    expect(value("B1")).toMatchObject({ type: "text", value: "HELLO" });
    expect(value("B2")).toMatchObject({ type: "text", value: "hello" });
  });

  test("CONCAT and CONCATENATE join text, numbers and booleans", () => {
    const { value } = calculate({
      B1: '=CONCAT("a","b",1)',
      B2: '=CONCATENATE("x","y")',
      B3: '=CONCATENATE("a",1,TRUE)',
    });
    expect(value("B1")).toMatchObject({ type: "text", value: "ab1" });
    expect(value("B2")).toMatchObject({ type: "text", value: "xy" });
    expect(value("B3")).toMatchObject({ type: "text", value: "a1TRUE" });
  });

  test("TEXT formats numbers, percentages and dates", () => {
    const { value } = calculate({
      A1: '=TEXT(1234.5,"#,##0.00")',
      A2: '=TEXT(0.256,"0%")',
      A3: '=TEXT(DATE(2024,3,5),"yyyy-mm-dd")',
    });
    expect(value("A1")).toMatchObject({ type: "text", value: "1,234.50" });
    expect(value("A2")).toMatchObject({ type: "text", value: "26%" });
    expect(value("A3")).toMatchObject({ type: "text", value: "2024-03-05" });
  });
});

describe("lookup functions", () => {
  const table = {
    A1: "a",
    B1: 1,
    A2: "b",
    B2: 2,
    A3: "c",
    B3: 3,
  };

  test("VLOOKUP with FALSE finds an exact match and returns the column value", () => {
    const { value } = calculate({
      ...table,
      C1: '=VLOOKUP("b",A1:B3,2,FALSE)',
    });
    expect(value("C1")).toMatchObject({ type: "number", value: 2 });
  });

  test("VLOOKUP with FALSE returns #N/A when no exact match exists", () => {
    const { value } = calculate({
      ...table,
      C1: '=VLOOKUP("z",A1:B3,2,FALSE)',
    });
    expect(value("C1")).toMatchObject({ type: "error", value: "#N/A" });
  });

  test("INDEX returns the value at a row and column position", () => {
    const { value } = calculate({
      ...table,
      C1: "=INDEX(B1:B3,3)",
      C2: "=INDEX(A1:B3,2,2)",
    });
    expect(value("C1")).toMatchObject({ type: "number", value: 3 });
    expect(value("C2")).toMatchObject({ type: "number", value: 2 });
  });

  test("MATCH with 0 returns the position of an exact match, or #N/A", () => {
    const { value } = calculate({
      ...table,
      C1: '=MATCH("b",A1:A3,0)',
      C2: '=MATCH("z",A1:A3,0)',
    });
    expect(value("C1")).toMatchObject({ type: "number", value: 2 });
    expect(value("C2")).toMatchObject({ type: "error", value: "#N/A" });
  });
});

describe("date functions with fixed inputs", () => {
  test("DATE builds the serial number for 29 February 2024, and YEAR, MONTH and DAY read it back", () => {
    const { value } = calculate({
      A1: "=DATE(2024,2,29)",
      A2: "=YEAR(A1)",
      A3: "=MONTH(A1)",
      A4: "=DAY(A1)",
    });
    expect(value("A1")).toMatchObject({ type: "number", value: 45351 });
    expect(value("A2")).toMatchObject({ type: "number", value: 2024 });
    expect(value("A3")).toMatchObject({ type: "number", value: 2 });
    expect(value("A4")).toMatchObject({ type: "number", value: 29 });
  });

  test("DATE rolls month overflow into the next year", () => {
    const { value } = calculate({
      A1: "=DATE(2023,14,1)",
      A2: "=YEAR(A1)",
      A3: "=MONTH(A1)",
      A4: "=DAY(A1)",
    });
    expect(value("A1")).toMatchObject({ type: "number", value: 45323 });
    expect(value("A2")).toMatchObject({ type: "number", value: 2024 });
    expect(value("A3")).toMatchObject({ type: "number", value: 2 });
    expect(value("A4")).toMatchObject({ type: "number", value: 1 });
  });

  test("EDATE clamps to the last day of a shorter month", () => {
    const { value } = calculate({
      A1: "=EDATE(DATE(2024,1,31),1)",
      A2: "=EDATE(DATE(2024,2,29),12)",
    });
    expect(value("A1")).toMatchObject({ type: "number", value: 45351 });
    expect(value("A2")).toMatchObject({ type: "number", value: 45716 });
  });

  test("EOMONTH returns the last day of the month offset from the start date", () => {
    const { value } = calculate({
      A1: "=EOMONTH(DATE(2023,1,15),1)",
      A2: "=EOMONTH(DATE(2024,1,31),0)",
      A3: "=DAY(EOMONTH(DATE(2023,2,10),0))",
    });
    expect(value("A1")).toMatchObject({ type: "number", value: 44985 });
    expect(value("A2")).toMatchObject({ type: "number", value: 45322 });
    expect(value("A3")).toMatchObject({ type: "number", value: 28 });
  });
});
