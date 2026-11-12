import { describe, expect, test } from "vitest";
import {
  buildCalculationEvidence,
  builtinFunctionNames,
  calculateWorkbookWithSettings,
  dateFunctionNames,
  databaseFunctionNames,
  daxFunctionNames,
  defaultCalculationSettings,
  estimateWorkbookMemoryBytes,
  expandWhatIfScenarios,
  getCurrentCalculatorRecord,
  matrixResultFunctionNames,
  normalizeCalculationSettings,
  refreshCalculationReport,
  resolveCalculationBudget,
  statisticalFunctionNames,
  supportedFinancialFunctionNames,
  supportedTextFunctionNames,
  textCleanupFunctionNames,
  validateWhatIfDefinition,
  whatIfLimits,
  withCalculationSettings,
  worksheetCapacityLimits,
} from "../../src/sdk/calculation.js";
import {
  createBlankWorkbook,
  getWorkbookCalculation,
} from "../../src/app/03-workbook-model.js";

/** A one-sheet workbook with the given raw cell contents. */
function workbookWith(cells) {
  const workbook = createBlankWorkbook("Calc");
  for (const [address, raw] of Object.entries(cells))
    workbook.sheets[0].cells[address] = { raw };
  return workbook;
}

const sheetIdOf = (workbook) => workbook.sheets[0].id;

/** A what-if request with two changing inputs and one formula result. */
function whatIfRequest(extra) {
  const sheetId = "sheet-1";
  return {
    name: "Plan",
    inputs: [
      { sheetId, address: "B2" },
      { sheetId, address: "C3" },
    ],
    outputs: [{ sheetId, address: "D1" }],
    ...extra,
  };
}

describe("normalizeCalculationSettings", () => {
  test("accepts the defaults and returns the same values", () => {
    expect(normalizeCalculationSettings(defaultCalculationSettings)).toEqual({
      mode: "automatic",
      iterative: false,
      maxIterations: 100,
      maxChange: 0.001,
      timeBudgetMs: 2500,
    });
  });

  test("keeps the optional memory and work budgets when they are given", () => {
    const normalized = normalizeCalculationSettings({
      ...defaultCalculationSettings,
      mode: "manual",
      memoryBudgetMiB: 64,
      workBudget: 5000,
    });
    expect(normalized.mode).toBe("manual");
    expect(normalized.memoryBudgetMiB).toBe(64);
    expect(normalized.workBudget).toBe(5000);
  });

  test("throws for an unknown key, an unknown mode and an iteration count above 1000", () => {
    expect(() =>
      normalizeCalculationSettings({ ...defaultCalculationSettings, extra: 1 }),
    ).toThrow("Invalid calculation settings.");
    expect(() =>
      normalizeCalculationSettings({
        ...defaultCalculationSettings,
        mode: "auto",
      }),
    ).toThrow("Invalid calculation settings.");
    expect(() =>
      normalizeCalculationSettings({
        ...defaultCalculationSettings,
        maxIterations: 1001,
      }),
    ).toThrow("Invalid calculation settings.");
  });

  test("throws when settings are missing or not an object", () => {
    expect(() => normalizeCalculationSettings()).toThrow(
      "Invalid calculation settings.",
    );
    expect(() => normalizeCalculationSettings([])).toThrow(
      "Invalid calculation settings.",
    );
  });
});

describe("withCalculationSettings", () => {
  test("returns a new workbook with the normalized settings and leaves the argument unchanged", () => {
    const workbook = workbookWith({ A1: 1 });
    const manual = withCalculationSettings(workbook, {
      ...defaultCalculationSettings,
      mode: "manual",
    });
    expect(manual).not.toBe(workbook);
    expect(manual.calculation.mode).toBe("manual");
    expect(manual.id).toBe(workbook.id);
    expect(manual.revision).toBe(workbook.revision);
    expect(workbook.calculation).toBeUndefined();
  });

  test("shares the sheets array with the argument", () => {
    const workbook = workbookWith({ A1: 1 });
    const copy = withCalculationSettings(workbook, {
      ...defaultCalculationSettings,
    });
    expect(copy.sheets).toBe(workbook.sheets);
  });

  test("throws for invalid settings", () => {
    expect(() =>
      withCalculationSettings(workbookWith({}), {
        ...defaultCalculationSettings,
        mode: "x",
      }),
    ).toThrow("Invalid calculation settings.");
  });
});

describe("defaultCalculationSettings", () => {
  test("holds the documented defaults and is frozen", () => {
    expect(defaultCalculationSettings).toEqual({
      mode: "automatic",
      iterative: false,
      maxIterations: 100,
      maxChange: 0.001,
      timeBudgetMs: 2500,
    });
    expect(Object.isFrozen(defaultCalculationSettings)).toBe(true);
  });

  test("is accepted unchanged by normalizeCalculationSettings", () => {
    expect(() =>
      normalizeCalculationSettings(defaultCalculationSettings),
    ).not.toThrow();
  });
});

describe("calculateWorkbookWithSettings", () => {
  test("applies filteredRowsBySheet, so SUBTOTAL skips a filtered row", () => {
    const workbook = workbookWith({
      A1: 1,
      A2: 2,
      A3: 4,
      B1: "=SUBTOTAL(9,A1:A3)",
    });
    const sheetId = sheetIdOf(workbook);
    expect(calculateWorkbookWithSettings(workbook)(sheetId, "B1")).toBe(7);
    const filtered = calculateWorkbookWithSettings(workbook, {
      filteredRowsBySheet: { [sheetId]: [{ start: 1, end: 1 }] },
    });
    expect(filtered(sheetId, "B1")).toBe(5);
  });

  test("throws for calculation settings, which it does not accept", () => {
    const workbook = workbookWith({ A1: 1 });
    expect(() =>
      calculateWorkbookWithSettings(workbook, { mode: "manual" }),
    ).toThrow("Calculation options support only filteredRowsBySheet.");
  });

  test("throws for a filter that names an unknown worksheet", () => {
    const workbook = workbookWith({ A1: 1 });
    expect(() =>
      calculateWorkbookWithSettings(workbook, {
        filteredRowsBySheet: { "no-such-sheet": [{ start: 0, end: 0 }] },
      }),
    ).toThrow("Filter context refers to an unknown worksheet ID.");
  });

  test("does not store a calculation for a plain workbook", () => {
    const workbook = workbookWith({ A1: 1 });
    calculateWorkbookWithSettings(workbook);
    expect(getCurrentCalculatorRecord(workbook)).toBeUndefined();
  });
});

describe("refreshCalculationReport", () => {
  test("recalculates, stores the new result and advances the run counter", () => {
    const workbook = workbookWith({ A1: 2, B1: "=A1*3" });
    const first = getWorkbookCalculation(workbook);
    const refreshed = refreshCalculationReport(workbook);
    expect(refreshed.report).not.toBe(first.report);
    expect(refreshed.report.run).toBe(first.report.run + 1);
    expect(getCurrentCalculatorRecord(workbook).report).toBe(refreshed.report);
  });

  test("turns a manual-mode pending result into a value", () => {
    const manual = withCalculationSettings(
      workbookWith({ A1: 2, B1: "=A1*3" }),
      {
        ...defaultCalculationSettings,
        mode: "manual",
      },
    );
    const sheetId = sheetIdOf(manual);
    expect(
      getWorkbookCalculation(manual).calculator.getTypedValue(sheetId, "B1"),
    ).toEqual({
      type: "error",
      value: "#PENDING!",
    });
    const refreshed = refreshCalculationReport(manual);
    expect(refreshed.calculator.getTypedValue(sheetId, "B1")).toEqual({
      type: "number",
      value: 6,
    });
  });
});

describe("getCurrentCalculatorRecord", () => {
  test("returns undefined before any calculation", () => {
    expect(getCurrentCalculatorRecord(workbookWith({ A1: 1 }))).toBeUndefined();
  });

  test("returns the same calculator and report that getWorkbookCalculation returned", () => {
    const workbook = workbookWith({ A1: 1 });
    const { calculator, report } = getWorkbookCalculation(workbook);
    const record = getCurrentCalculatorRecord(workbook);
    expect(record.report).toBe(report);
    expect(record.calculator).toBe(calculator);
  });

  test("returns undefined for a new workbook object built from the old one", () => {
    const workbook = workbookWith({ A1: 1 });
    getWorkbookCalculation(workbook);
    const edited = {
      ...workbook,
      sheets: [{ ...workbook.sheets[0], cells: { A1: { raw: 5 } } }],
    };
    expect(getCurrentCalculatorRecord(edited)).toBeUndefined();
  });
});

describe("buildCalculationEvidence", () => {
  test("returns a JSON string that holds the report and the evidence boundaries", () => {
    const workbook = workbookWith({ A1: 1, B1: "=A1+1" });
    const { report } = getWorkbookCalculation(workbook);
    const evidence = JSON.parse(buildCalculationEvidence(workbook, report));
    expect(evidence.schemaVersion).toBe(1);
    expect(evidence.report.workbookId).toBe(workbook.id);
    expect(evidence.boundaries.length).toBeGreaterThan(0);
    expect(evidence.imported).toBeNull();
  });

  test("throws when the report is no longer the current one", () => {
    const workbook = workbookWith({ A1: 1, B1: "=A1+1" });
    const { report } = getWorkbookCalculation(workbook);
    refreshCalculationReport(workbook);
    expect(() => buildCalculationEvidence(workbook, report)).toThrow(
      "The workbook or calculation run changed.",
    );
  });
});

describe("function name catalogues", () => {
  test("builtinFunctionNames lists uppercase names without duplicates, including SUM and XLOOKUP", () => {
    expect(builtinFunctionNames).toContain("SUM");
    expect(builtinFunctionNames).toContain("XLOOKUP");
    expect(new Set(builtinFunctionNames).size).toBe(
      builtinFunctionNames.length,
    );
    expect(
      builtinFunctionNames.every((name) => name === name.toUpperCase()),
    ).toBe(true);
  });

  test("statisticalFunctionNames is a subset of builtinFunctionNames", () => {
    expect(statisticalFunctionNames).toContain("AVERAGE");
    expect(statisticalFunctionNames).toContain("MEDIAN");
    for (const name of statisticalFunctionNames)
      expect(builtinFunctionNames).toContain(name);
  });

  test("databaseFunctionNames holds the D-functions, including DSUM and DGET", () => {
    expect(databaseFunctionNames).toContain("DSUM");
    expect(databaseFunctionNames).toContain("DGET");
    expect(databaseFunctionNames.every((name) => name.startsWith("D"))).toBe(
      true,
    );
  });

  test("matrixResultFunctionNames includes SEQUENCE and FILTER but not SUM", () => {
    expect(matrixResultFunctionNames).toContain("SEQUENCE");
    expect(matrixResultFunctionNames).toContain("FILTER");
    expect(matrixResultFunctionNames).not.toContain("SUM");
  });

  test("daxFunctionNames maps names to [minimum, maximum] argument counts", () => {
    expect(daxFunctionNames.BLANK).toEqual([0, 0]);
    expect(daxFunctionNames.IF).toEqual([2, 3]);
    expect(Object.isFrozen(daxFunctionNames)).toBe(true);
  });

  test("textCleanupFunctionNames includes TRIM and CLEAN", () => {
    expect(textCleanupFunctionNames).toContain("TRIM");
    expect(textCleanupFunctionNames).toContain("CLEAN");
  });

  test("dateFunctionNames includes DATE and EDATE", () => {
    expect(dateFunctionNames).toContain("DATE");
    expect(dateFunctionNames).toContain("EDATE");
  });

  test("supportedFinancialFunctionNames includes PMT and NPV", () => {
    expect(supportedFinancialFunctionNames).toContain("PMT");
    expect(supportedFinancialFunctionNames).toContain("NPV");
  });

  test("supportedTextFunctionNames includes LEFT and TEXTJOIN", () => {
    expect(supportedTextFunctionNames).toContain("LEFT");
    expect(supportedTextFunctionNames).toContain("TEXTJOIN");
  });
});

describe("resolveCalculationBudget", () => {
  test("uses 32 MiB and 1,000,000 work units when no budget is given", () => {
    expect(resolveCalculationBudget()).toEqual({
      memoryBudgetBytes: 32 * 1024 * 1024,
      workBudget: 1000000,
    });
    expect(resolveCalculationBudget({})).toEqual({
      memoryBudgetBytes: 32 * 1024 * 1024,
      workBudget: 1000000,
    });
  });

  test("converts a memory budget in MiB to bytes and keeps the work budget", () => {
    expect(
      resolveCalculationBudget({ memoryBudgetMiB: 64, workBudget: 5000 }),
    ).toEqual({
      memoryBudgetBytes: 64 * 1024 * 1024,
      workBudget: 5000,
    });
  });
});

describe("estimateWorkbookMemoryBytes", () => {
  test("counts more bytes for a longer text cell", () => {
    const small = workbookWith({ A1: "x" });
    const large = workbookWith({ A1: "x".repeat(1000) });
    const budget = 33554432;
    const smallBytes = estimateWorkbookMemoryBytes(small, budget);
    const largeBytes = estimateWorkbookMemoryBytes(large, budget);
    expect(smallBytes).toBeGreaterThan(0);
    // Text is counted at two bytes per UTF-16 unit, so 999 extra characters add 1998 bytes.
    expect(largeBytes - smallBytes).toBe(2 * (1000 - 1));
  });

  test("returns a value above the budget when the budget is too small", () => {
    const workbook = workbookWith({ A1: "x".repeat(1000) });
    expect(estimateWorkbookMemoryBytes(workbook, 100)).toBeGreaterThan(100);
  });

  test("throws when the argument is not a workbook", () => {
    expect(() => estimateWorkbookMemoryBytes({}, 1000)).toThrow(TypeError);
  });
});

describe("worksheetCapacityLimits and whatIfLimits", () => {
  test("worksheetCapacityLimits gives the hidden-interval and frozen-pane caps", () => {
    expect(worksheetCapacityLimits.maxIntervalsPerAxis).toBe(10000);
    expect(worksheetCapacityLimits.maxFrozenVisibleRows).toBe(100);
  });

  test("whatIfLimits gives the input, scenario and trial caps", () => {
    expect(whatIfLimits.inputs).toBe(32);
    expect(whatIfLimits.scenarios).toBe(100);
    expect(whatIfLimits.trials).toBe(400);
  });
});

describe("validateWhatIfDefinition", () => {
  test("accepts a scenario definition, trims its name and normalizes addresses", () => {
    const definition = validateWhatIfDefinition({
      ...whatIfRequest({ name: "  Base  " }),
      inputs: [
        { sheetId: "sheet-1", address: "b2" },
        { sheetId: "sheet-1", address: "C3" },
      ],
      kind: "scenarios",
      scenarios: [
        { name: "Low", values: [1, 2] },
        { name: "High", values: [3, 4] },
      ],
    });
    expect(definition.name).toBe("Base");
    expect(definition.inputs[0]).toEqual({ sheetId: "sheet-1", address: "B2" });
    expect(definition.scenarios).toHaveLength(2);
  });

  test("accepts a two-variable table within the 400 combination limit", () => {
    const definition = validateWhatIfDefinition({
      ...whatIfRequest({}),
      kind: "two-variable",
      rowValues: [1, 2],
      columnValues: [3, 4, 5],
    });
    expect(definition.kind).toBe("two-variable");
  });

  test("throws for scenario names that differ only by case", () => {
    expect(() =>
      validateWhatIfDefinition({
        ...whatIfRequest({}),
        kind: "scenarios",
        scenarios: [
          { name: "low", values: [1, 2] },
          { name: "LOW", values: [3, 4] },
        ],
      }),
    ).toThrow("Scenario names must be unique, ignoring case.");
  });

  test("throws when an output is also an input", () => {
    expect(() =>
      validateWhatIfDefinition({
        name: "Loop",
        kind: "scenarios",
        inputs: [{ sheetId: "sheet-1", address: "D1" }],
        outputs: [{ sheetId: "sheet-1", address: "D1" }],
        scenarios: [{ name: "a", values: [1] }],
      }),
    ).toThrow("Changing inputs and formula results must be different cells.");
  });

  test("throws for a two-variable table with more than 400 combinations", () => {
    expect(() =>
      validateWhatIfDefinition({
        ...whatIfRequest({}),
        kind: "two-variable",
        rowValues: Array(21).fill(1),
        columnValues: Array(20).fill(1),
      }),
    ).toThrow("A sensitivity table supports at most 400 combinations.");
  });
});

describe("expandWhatIfScenarios", () => {
  test("returns one case per scenario with its input values", () => {
    const definition = validateWhatIfDefinition({
      ...whatIfRequest({}),
      kind: "scenarios",
      scenarios: [{ name: "Low", values: [1, 2] }],
    });
    expect(expandWhatIfScenarios(definition)).toEqual([
      { name: "Low", inputs: [1, 2] },
    ]);
  });

  test("names each one-variable case by its value", () => {
    const definition = validateWhatIfDefinition({
      ...whatIfRequest({ inputs: [{ sheetId: "sheet-1", address: "B2" }] }),
      kind: "one-variable",
      orientation: "column",
      values: [0.5, 1],
    });
    expect(expandWhatIfScenarios(definition)).toEqual([
      { name: "0.5", inputs: [0.5] },
      { name: "1", inputs: [1] },
    ]);
  });

  test("produces every row and column pair for a two-variable table", () => {
    const definition = validateWhatIfDefinition({
      ...whatIfRequest({}),
      kind: "two-variable",
      rowValues: [1, 2],
      columnValues: [10],
    });
    const cases = expandWhatIfScenarios(definition);
    expect(cases).toEqual([
      { name: "1 / 10", inputs: [1, 10] },
      { name: "2 / 10", inputs: [2, 10] },
    ]);
  });

  test("throws a TypeError for a request it cannot read, because it does not validate", () => {
    expect(() => expandWhatIfScenarios({ kind: "nope" })).toThrow(TypeError);
  });
});
