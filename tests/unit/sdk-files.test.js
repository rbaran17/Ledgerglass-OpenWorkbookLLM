import { describe, expect, test } from "vitest";
import {
  createBlankWorkbook,
  getWorkbookCalculation,
  normalizeWorkbookInput,
} from "../../src/app/03-workbook-model.js";
import {
  FileProcessingError,
  delimitedImportDefaults,
  encodeNativeWorkbook,
  exportWorkbookToXlsx,
  exportWorksheetToCsv,
  importDelimitedFile,
  importNativeJsonFile,
  importWorkbookFile,
} from "../../src/sdk/files.js";

function budgetWorkbook() {
  const workbook = createBlankWorkbook("Budget");
  Object.assign(workbook.sheets[0].cells, {
    A1: { raw: "Item" },
    B1: { raw: "Amount" },
    A2: { raw: "Rent" },
    B2: { raw: 1200 },
    A3: { raw: "Food" },
    B3: { raw: 450.5 },
    A4: { raw: "Total" },
    B4: { raw: "=SUM(B2:B3)" },
  });
  return workbook;
}
const valueOf = (workbook, address) =>
  getWorkbookCalculation(workbook).calculator.getTypedValue(
    workbook.sheets[0].id,
    address,
  );
const xlsxType =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

describe("XLSX", () => {
  test("export returns a blob, warnings and a report", async () => {
    const result = await exportWorkbookToXlsx(budgetWorkbook());
    expect(result.blob).toBeInstanceOf(Blob);
    expect(result.blob.type).toBe(xlsxType);
    expect(result.blob.size).toBeGreaterThan(1000);
    expect(result.warnings.every((item) => typeof item === "string")).toBe(
      true,
    );
    expect(result.report).toBeTypeOf("object");
  });

  test("an exported workbook imports with values and formulas intact", async () => {
    const { blob } = await exportWorkbookToXlsx(budgetWorkbook());
    const imported = await importWorkbookFile(new File([blob], "budget.xlsx"));
    const cells = imported.workbook.sheets[0].cells;
    expect(cells.A2.raw).toBe("Rent");
    expect(cells.B3.raw).toBe(450.5);
    expect(cells.B4.raw).toBe("=SUM(B2:B3)");
    expect(valueOf(imported.workbook, "B4")).toMatchObject({
      type: "number",
      value: 1650.5,
    });
    expect(Array.isArray(imported.warnings)).toBe(true);
  });

  test("a file that is not an XLSX archive is refused with FILE_INVALID", async () => {
    const attempt = importWorkbookFile(new File(["not a zip"], "bad.xlsx"));
    await expect(attempt).rejects.toBeInstanceOf(FileProcessingError);
    await expect(attempt).rejects.toMatchObject({ code: "FILE_INVALID" });
  });

  test("an extension that is not accepted is refused with FILE_TYPE", async () => {
    await expect(
      importWorkbookFile(new File(["x"], "tool.exe")),
    ).rejects.toMatchObject({ name: "FileProcessingError", code: "FILE_TYPE" });
  });

  test("an import that is already cancelled is refused with FILE_CANCELLED", async () => {
    const { blob } = await exportWorkbookToXlsx(budgetWorkbook());
    await expect(
      importWorkbookFile(new File([blob], "budget.xlsx"), {
        signal: AbortSignal.abort(),
      }),
    ).rejects.toMatchObject({ code: "FILE_CANCELLED" });
  });
});

describe("delimited text", () => {
  test("export writes calculated values with CRLF line ends", () => {
    const workbook = budgetWorkbook();
    const result = exportWorksheetToCsv(workbook, workbook.sheets[0].id);
    expect(result.text).toBe(
      "Item,Amount\r\nRent,1200\r\nFood,450.5\r\nTotal,1650.5",
    );
    expect(new TextDecoder().decode(result.bytes)).toBe(result.text);
  });

  test("export honours the delimiter and a range", () => {
    const workbook = budgetWorkbook();
    const result = exportWorksheetToCsv(workbook, workbook.sheets[0].id, {
      delimiter: ";",
      range: "A1:B2",
    });
    expect(result.text).toBe("Item;Amount\r\nRent;1200");
  });

  test("export throws for a sheet that does not exist", () => {
    expect(() => exportWorksheetToCsv(budgetWorkbook(), "missing")).toThrow(
      "Worksheet not found.",
    );
  });

  test("import types numbers and reports what it read", async () => {
    const result = await importDelimitedFile(
      new File(["name,amount\nRent,1200\nFood,450.5\n"], "budget.csv"),
    );
    const cells = result.workbook.sheets[0].cells;
    expect(cells.A1.raw).toBe("name");
    expect(cells.B2.raw).toBe(1200);
    expect(cells.B3.raw).toBe(450.5);
    expect(result.report).toMatchObject({
      direction: "import",
      fileName: "budget.csv",
      records: 3,
      columns: 2,
    });
  });

  test("a decimal comma is read for columns typed as numbers", async () => {
    const file = () => new File(["a;b\n1,5;2\n"], "values.csv");
    const typed = await importDelimitedFile(file(), {
      delimiter: ";",
      decimal: ",",
      columns: { 0: "number" },
    });
    expect(typed.workbook.sheets[0].cells.A2.raw).toBe(1.5);
    const untyped = await importDelimitedFile(file(), {
      delimiter: ";",
      decimal: ",",
    });
    expect(untyped.workbook.sheets[0].cells.A2.raw).toBe("1,5");
  });

  test("text that looks like a formula is kept as text, not calculated", async () => {
    const result = await importDelimitedFile(
      new File(["value\n=1+1\n"], "formula.csv"),
    );
    expect(result.workbook.sheets[0].cells.A2).toMatchObject({
      raw: "=1+1",
      literal: true,
    });
  });

  test("an unsupported setting is refused", async () => {
    await expect(
      importDelimitedFile(new File(["a\n1\n"], "x.csv"), { decimal: ";" }),
    ).rejects.toBeInstanceOf(FileProcessingError);
  });

  test("the defaults detect the delimiter and expect a header", () => {
    expect(delimitedImportDefaults).toMatchObject({
      delimiter: "auto",
      hasHeader: true,
      decimal: ".",
    });
  });
});

describe("native JSON", () => {
  test("encode gives JSON text that imports back to the same workbook", async () => {
    const workbook = budgetWorkbook();
    const text = encodeNativeWorkbook(workbook);
    expect(JSON.parse(text)).toMatchObject({
      schemaVersion: 1,
      id: workbook.id,
      name: "Budget",
    });
    const imported = await importNativeJsonFile(
      new File([text], "budget.json"),
    );
    expect(imported.id).toBe(workbook.id);
    expect(imported.sheets[0].cells).toEqual(workbook.sheets[0].cells);
  });

  test("the import parses but does not validate; normalizeWorkbookInput does", async () => {
    const parsed = await importNativeJsonFile(
      new File(['{"hello":1}'], "other.json"),
    );
    expect(parsed).toEqual({ hello: 1 });
    expect(() => normalizeWorkbookInput(parsed)).toThrow();
  });

  test("text that is not JSON is refused", async () => {
    await expect(
      importNativeJsonFile(new File(["{not json"], "broken.json")),
    ).rejects.toBeInstanceOf(FileProcessingError);
  });
});
