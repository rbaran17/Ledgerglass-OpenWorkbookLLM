/** Reading calculated values in bulk, for scripts and servers. */
import { getWorkbookCalculation } from "../app/03-workbook-model.js";

/**
 * Calculates the workbook and returns every filled cell's result.
 *
 * @returns `{ report, sheets }` where `sheets[sheetName][address]` is `{ type, value }` and
 *   `report` is the calculation report. Errors appear as `{ type: "error", value: "#DIV/0!" }`.
 */
export function calculateWorkbookValues(workbook) {
  const { calculator, report } = getWorkbookCalculation(workbook);
  const sheets = {};
  for (const sheet of workbook.sheets) {
    const values = {};
    for (const address of Object.keys(sheet.cells)) {
      const result = calculator.getTypedValue(sheet.id, address);
      values[address] = { type: result.type, value: result.value };
    }
    sheets[sheet.name] = values;
  }
  return { report, sheets };
}
