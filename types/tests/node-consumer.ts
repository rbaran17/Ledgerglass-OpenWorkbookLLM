/** Uses the Node entry the way a script would; checked by `npm run test:types`. */
import {
  FileProcessingError,
  addWorksheet,
  calculateWorkbookValues,
  createBlankWorkbook,
  exportWorksheetToCsv,
  readWorkbookFile,
  writeWorkbookFile,
  type Workbook,
} from "@ledgerglass/workbook/node";

export async function roundTrip(path: string): Promise<number> {
  const { workbook, warnings } = await readWorkbookFile(path, {
    delimited: { decimal: ",", columns: { 1: "number" } },
  });
  const widened: Workbook = addWorksheet(workbook, "Notes");
  const { sheets, report } = calculateWorkbookValues(widened);
  const first = sheets[widened.sheets[0].name].A1;
  if (first.type === "number") first.value.toFixed(2);
  const written = await writeWorkbookFile(widened, "out.xlsx");
  const csv: string = exportWorksheetToCsv(widened, widened.sheets[0].id).text;
  return warnings.length + written.bytes + report.formulaCount + csv.length;
}

export function describe(error: unknown): string {
  return error instanceof FileProcessingError ? error.code : "other";
}

import { AuditWorkbook } from "@ledgerglass/workbook/node";
// @ts-expect-error the interface is not part of the Node entry: only its type is
void AuditWorkbook;
// @ts-expect-error a path is required
void readWorkbookFile();
