import { expect, test } from "vitest";
import {
  createBlankWorkbook,
  getWorkbookCalculation,
} from "../../src/app/03-workbook-model.js";

test("a blank workbook calculates a simple formula", () => {
  const workbook = createBlankWorkbook("Unit");
  const sheet = workbook.sheets[0];
  sheet.cells.A1 = { raw: 2 };
  sheet.cells.A2 = { raw: "=A1*21" };
  const { calculator } = getWorkbookCalculation(workbook);
  expect(calculator.getTypedValue(sheet.id, "A2")).toMatchObject({ value: 42 });
});
