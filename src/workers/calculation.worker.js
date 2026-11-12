import {
  deepFreezeValue,
  estimateValueFootprintBytes,
  estimateWorkbookMemoryBytes,
  resolveCalculationBudget,
  setExtensionCacheEntries,
} from "../app/01-core-schemas-and-hashing.js";

import {
  carryCalculatorRecordForward,
  getCalculationInputSignature,
  zie,
} from "../app/03-workbook-model/05-consolidation-and-sources.js";

import {
  getWorkbookCalculation,
  refreshCalculationReport,
} from "../app/03-workbook-model/09-calculation-engine.js";

function GR(e, t) {
  let r = t.calculator,
    o = [];
  for (let n of e.sheets)
    try {
      for (let i of r.getSpillEntries(n.id))
        o.push({
          ...i,
          sheetId: n.id,
          result: r.getTypedValue(n.id, i.address),
        });
    } catch {
      if (!r.getOperationalErrors().includes("#LIMIT!"))
        throw Error("Could not capture spill results.");
    }
  return {
    inputSignature: getCalculationInputSignature(e),
    report: t.report,
    spills: o,
    spillMetadata: t.report.cells.map((n) => ({
      sheetId: n.sheetId,
      address: n.address,
      ...r.peekSpill(n.sheetId, n.address),
    })),
    references: r.getReferenceDiagnostics(),
    names: r.getNamedFormulaDiagnostics(),
    financial: r.getFinancialDiagnostics(),
    operationalErrors: r.getOperationalErrors(),
  };
}

function HR() {
  let e;
  return (t) => {
    let r = t.workbook;
    try {
      if (t.baseRevision !== undefined) {
        if (!e || e.id !== r.id || e.revision !== t.baseRevision)
          throw Error("Stale calculation delta");
        r = {
          ...r,
          sheets: r.sheets.map((a) => {
            let l = e.sheets.find((u) => u.id === a.id),
              s = t.removedCells?.[a.id] ?? [];
            if (!l) return a;
            if (!Object.keys(a.cells).length && !s.length)
              return { ...a, cells: l.cells };
            let c = { ...l.cells, ...a.cells };
            for (let u of s) delete c[u];
            return { ...a, cells: c };
          }),
        };
      }
      let { memoryBudgetBytes } = resolveCalculationBudget(r.calculation);
      if (estimateWorkbookMemoryBytes(r, memoryBudgetBytes) > memoryBudgetBytes)
        throw Error(
          "#LIMIT! Calculation inputs exceed the working-data budget.",
        );
      (deepFreezeValue(r),
        setExtensionCacheEntries(r, t.extensionAvailability ?? []),
        e && (zie(e, r), carryCalculatorRecordForward(e, r)));
      let n = t.force ? refreshCalculationReport(r) : getWorkbookCalculation(r),
        i = GR(r, n);
      if (estimateValueFootprintBytes(i, memoryBudgetBytes) > memoryBudgetBytes)
        throw Error(
          "#LIMIT! Calculation result payload exceeds the working-data budget.",
        );
      return ((e = r), { id: t.id, result: i });
    } catch (o) {
      return (
        (e = undefined),
        {
          id: t.id,
          error:
            o instanceof Error && o.message.startsWith("#LIMIT!")
              ? o.message
              : "Calculation worker failed; no result was published.",
        }
      );
    }
  };
}
var qD = HR(),
  KR = self;
KR.onmessage = ({ data }) => KR.postMessage(qD(data));
