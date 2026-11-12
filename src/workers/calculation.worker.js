import {
  deepFreezeValue,
  estimateValueFootprintBytes,
  estimateWorkbookMemoryBytes,
  resolveCalculationBudget,
  setExtensionCacheEntries,
} from "../app/01-schemas-date-functions-and-validation.js";

import {
  carryCalculatorRecordForward,
  getCalculationInputSignature,
  zie,
} from "../app/03-workbook-model/05-formula-engine-and-name-operations.js";

import {
  getWorkbookCalculation,
  refreshCalculationReport,
} from "../app/03-workbook-model/09-calculation-engine.js";

function GR(workbook, calculation) {
  let calculator = calculation.calculator,
    spills = [];
  for (let sheet of workbook.sheets)
    try {
      for (let spillEntry of calculator.getSpillEntries(sheet.id))
        spills.push({
          ...spillEntry,
          sheetId: sheet.id,
          result: calculator.getTypedValue(sheet.id, spillEntry.address),
        });
    } catch {
      if (!calculator.getOperationalErrors().includes("#LIMIT!"))
        throw Error("Could not capture spill results.");
    }
  return {
    inputSignature: getCalculationInputSignature(workbook),
    report: calculation.report,
    spills: spills,
    spillMetadata: calculation.report.cells.map((cell) => ({
      sheetId: cell.sheetId,
      address: cell.address,
      ...calculator.peekSpill(cell.sheetId, cell.address),
    })),
    references: calculator.getReferenceDiagnostics(),
    names: calculator.getNamedFormulaDiagnostics(),
    financial: calculator.getFinancialDiagnostics(),
    operationalErrors: calculator.getOperationalErrors(),
  };
}

function HR() {
  let previousWorkbook;
  return (request) => {
    let workbook = request.workbook;
    try {
      if (request.baseRevision !== undefined) {
        if (
          !previousWorkbook ||
          previousWorkbook.id !== workbook.id ||
          previousWorkbook.revision !== request.baseRevision
        )
          throw Error("Stale calculation delta");
        workbook = {
          ...workbook,
          sheets: workbook.sheets.map((sheet) => {
            let matchingSheet = previousWorkbook.sheets.find(
                (previousSheet) => previousSheet.id === sheet.id,
              ),
              removedCellAddresses = request.removedCells?.[sheet.id] ?? [];
            if (!matchingSheet) return sheet;
            if (
              !Object.keys(sheet.cells).length &&
              !removedCellAddresses.length
            )
              return { ...sheet, cells: matchingSheet.cells };
            let cells = { ...matchingSheet.cells, ...sheet.cells };
            for (let cellAddress of removedCellAddresses)
              delete cells[cellAddress];
            return { ...sheet, cells: cells };
          }),
        };
      }
      let { memoryBudgetBytes } = resolveCalculationBudget(
        workbook.calculation,
      );
      if (
        estimateWorkbookMemoryBytes(workbook, memoryBudgetBytes) >
        memoryBudgetBytes
      )
        throw Error(
          "#LIMIT! Calculation inputs exceed the working-data budget.",
        );
      (deepFreezeValue(workbook),
        setExtensionCacheEntries(workbook, request.extensionAvailability ?? []),
        previousWorkbook &&
          (zie(previousWorkbook, workbook),
          carryCalculatorRecordForward(previousWorkbook, workbook)));
      let calculation = request.force
          ? refreshCalculationReport(workbook)
          : getWorkbookCalculation(workbook),
        result = GR(workbook, calculation);
      if (
        estimateValueFootprintBytes(result, memoryBudgetBytes) >
        memoryBudgetBytes
      )
        throw Error(
          "#LIMIT! Calculation result payload exceeds the working-data budget.",
        );
      return (
        (previousWorkbook = workbook),
        { id: request.id, result: result }
      );
    } catch (error) {
      return (
        (previousWorkbook = undefined),
        {
          id: request.id,
          error:
            error instanceof Error && error.message.startsWith("#LIMIT!")
              ? error.message
              : "Calculation worker failed; no result was published.",
        }
      );
    }
  };
}
var qD = HR(),
  KR = self;
KR.onmessage = ({ data }) => KR.postMessage(qD(data));
