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

function GR(workbook, t) {
  let calculator = t.calculator,
    spills = [];
  for (let sheet of workbook.sheets)
    try {
      for (let i of calculator.getSpillEntries(sheet.id))
        spills.push({
          ...i,
          sheetId: sheet.id,
          result: calculator.getTypedValue(sheet.id, i.address),
        });
    } catch {
      if (!calculator.getOperationalErrors().includes("#LIMIT!"))
        throw Error("Could not capture spill results.");
    }
  return {
    inputSignature: getCalculationInputSignature(workbook),
    report: t.report,
    spills: spills,
    spillMetadata: t.report.cells.map((cell) => ({
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
  let e;
  return (t) => {
    let workbook = t.workbook;
    try {
      if (t.baseRevision !== undefined) {
        if (!e || e.id !== workbook.id || e.revision !== t.baseRevision)
          throw Error("Stale calculation delta");
        workbook = {
          ...workbook,
          sheets: workbook.sheets.map((sheet) => {
            let l = e.sheets.find((u) => u.id === sheet.id),
              s = t.removedCells?.[sheet.id] ?? [];
            if (!l) return sheet;
            if (!Object.keys(sheet.cells).length && !s.length)
              return { ...sheet, cells: l.cells };
            let cells = { ...l.cells, ...sheet.cells };
            for (let u of s) delete cells[u];
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
        setExtensionCacheEntries(workbook, t.extensionAvailability ?? []),
        e && (zie(e, workbook), carryCalculatorRecordForward(e, workbook)));
      let n = t.force
          ? refreshCalculationReport(workbook)
          : getWorkbookCalculation(workbook),
        i = GR(workbook, n);
      if (estimateValueFootprintBytes(i, memoryBudgetBytes) > memoryBudgetBytes)
        throw Error(
          "#LIMIT! Calculation result payload exceeds the working-data budget.",
        );
      return ((e = workbook), { id: t.id, result: i });
    } catch (error) {
      return (
        (e = undefined),
        {
          id: t.id,
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
