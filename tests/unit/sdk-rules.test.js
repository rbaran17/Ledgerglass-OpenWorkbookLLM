import { describe, expect, test } from "vitest";
import {
  addReviewThreadReply,
  applyDefinedNameOperation,
  assertConditionalFormatPreviewCurrent,
  assertNamePreviewCurrent,
  assertNoStopValidationIssues,
  checkDataValidationBatch,
  conditionalRuleKinds,
  DataValidationError,
  evaluateConditionalFormatRules,
  findDataValidationsInRange,
  normalizeConditionalRules,
  normalizeDataValidationRules,
  resolveDefinedNameValue,
  resolveValidationInspection,
  reviewThreadLimits,
  setCellHyperlink,
  setReviewThreadResolved,
  sortConditionalRulesByPriority,
  summarizeValidationInspection,
  updateConditionalFormatRules,
  validateDataValidationRule,
  validateFormulaWatches,
  validateHyperlinkRecord,
} from "../../src/sdk/rules.js";
import {
  createBlankWorkbook,
  getWorkbookCalculation,
} from "../../src/app/03-workbook-model.js";

const STAMP = "2026-01-02T09:00:00.000Z";

function numberRule(overrides = {}) {
  return {
    id: "num",
    range: "A1:A3",
    kind: "number",
    operator: "between",
    min: 1,
    max: 10,
    list: [],
    allowBlank: true,
    message: "Enter 1 to 10",
    ...overrides,
  };
}

function conditionalRule(overrides = {}) {
  return {
    id: "r1",
    range: "A1:A4",
    kind: "greater",
    value: "10",
    second: "",
    fill: "#ffcccc",
    color: "#000000",
    ...overrides,
  };
}

/** A one-sheet workbook with A1=5, A2=50, A3 blank, B1 a formula. */
function sampleWorkbook() {
  const workbook = createBlankWorkbook("Rules");
  const sheet = workbook.sheets[0];
  sheet.cells.A1 = { raw: 5 };
  sheet.cells.A2 = { raw: 50 };
  sheet.cells.B1 = { raw: "=A1*2" };
  return workbook;
}

describe("data validation rules", () => {
  test("normalizeDataValidationRules upper-cases and orders the range", () => {
    const input = [numberRule({ range: "b2:a1" })];
    const [rule] = normalizeDataValidationRules(input);
    expect(rule.range).toBe("A1:B2");
    expect(input[0].range).toBe("b2:a1");
  });

  test("normalizeDataValidationRules rejects a rule without a message", () => {
    expect(() =>
      normalizeDataValidationRules([numberRule({ message: undefined })]),
    ).toThrow("Invalid validation error message.");
  });

  test("normalizeDataValidationRules rejects repeated IDs and empty lists", () => {
    expect(() =>
      normalizeDataValidationRules([numberRule(), numberRule()]),
    ).toThrow("Validation rule IDs must be nonempty and unique.");
    expect(() =>
      normalizeDataValidationRules([
        numberRule({ id: "l", kind: "list", operator: "equal", list: [] }),
      ]),
    ).toThrow("A list rule needs nonempty values.");
  });

  test("validateDataValidationRule applies bounds and blank handling", () => {
    const rule = normalizeDataValidationRules([numberRule()])[0];
    expect(validateDataValidationRule(rule, 5)).toBe(true);
    expect(validateDataValidationRule(rule, 11)).toBe(false);
    expect(validateDataValidationRule(rule, "")).toBe(true);
    expect(validateDataValidationRule(rule, "abc")).toBe(false);
  });

  test("validateDataValidationRule compares list items without case", () => {
    const rule = normalizeDataValidationRules([
      numberRule({
        id: "l",
        kind: "list",
        operator: "equal",
        min: 0,
        max: 0,
        list: ["Yes", "No"],
        allowBlank: false,
      }),
    ])[0];
    expect(validateDataValidationRule(rule, "yes")).toBe(true);
    expect(validateDataValidationRule(rule, "Maybe")).toBe(false);
  });

  test("validateDataValidationRule accepts whole numbers only for integer rules", () => {
    const rule = normalizeDataValidationRules([
      numberRule({
        id: "i",
        kind: "integer",
        operator: "greater",
        min: 0,
        max: 0,
      }),
    ])[0];
    expect(validateDataValidationRule(rule, 2)).toBe(true);
    expect(validateDataValidationRule(rule, 2.5)).toBe(false);
  });

  test("validateDataValidationRule returns false for custom rules", () => {
    const rule = normalizeDataValidationRules([
      numberRule({
        id: "c",
        kind: "custom",
        operator: "equal",
        min: 0,
        max: 0,
        formula: "=B1>0",
      }),
    ])[0];
    expect(validateDataValidationRule(rule, 3)).toBe(false);
  });

  test("findDataValidationsInRange returns the rules covering a cell", () => {
    const workbook = sampleWorkbook();
    const sheetId = workbook.sheets[0].id;
    workbook.sheets[0].validations = normalizeDataValidationRules([
      numberRule(),
    ]);
    expect(
      findDataValidationsInRange(workbook, sheetId, "A2").map(
        (rule) => rule.id,
      ),
    ).toEqual(["num"]);
    expect(
      findDataValidationsInRange(workbook, sheetId, "a2").map(
        (rule) => rule.id,
      ),
    ).toEqual(["num"]);
  });

  test("findDataValidationsInRange returns an empty array outside every range", () => {
    const workbook = sampleWorkbook();
    const sheetId = workbook.sheets[0].id;
    workbook.sheets[0].validations = normalizeDataValidationRules([
      numberRule(),
    ]);
    expect(findDataValidationsInRange(workbook, sheetId, "C9")).toEqual([]);
  });

  test("checkDataValidationBatch returns 0 for valid values and for no targets", () => {
    const workbook = sampleWorkbook();
    const sheetId = workbook.sheets[0].id;
    workbook.sheets[0].validations = normalizeDataValidationRules([
      numberRule(),
    ]);
    expect(checkDataValidationBatch(workbook, [])).toBe(0);
    expect(
      checkDataValidationBatch(workbook, [{ sheetId, address: "A1" }]),
    ).toBe(0);
  });

  test("checkDataValidationBatch throws DataValidationError for a stop violation", () => {
    const workbook = sampleWorkbook();
    const sheetId = workbook.sheets[0].id;
    workbook.sheets[0].validations = normalizeDataValidationRules([
      numberRule(),
    ]);
    let thrown;
    try {
      checkDataValidationBatch(workbook, [{ sheetId, address: "A2" }]);
    } catch (error) {
      thrown = error;
    }
    expect(thrown).toBeInstanceOf(DataValidationError);
    expect(thrown.issues).toHaveLength(1);
    expect(thrown.issues[0]).toMatchObject({
      address: "A2",
      style: "stop",
      ruleId: "num",
      value: 50,
    });
  });

  test("a warning-only batch proceeds once it is approved", () => {
    const workbook = sampleWorkbook();
    const sheetId = workbook.sheets[0].id;
    workbook.sheets[0].validations = normalizeDataValidationRules([
      numberRule({
        id: "warn",
        range: "B1:B1",
        operator: "lessEqual",
        min: 5,
        max: 0,
        allowBlank: false,
        message: "Above 5",
        errorStyle: "warning",
      }),
    ]);
    const targets = [{ sheetId, address: "B1" }];
    let thrown;
    try {
      checkDataValidationBatch(workbook, targets);
    } catch (error) {
      thrown = error;
    }
    expect(thrown).toBeInstanceOf(DataValidationError);
    const approval = assertNoStopValidationIssues(thrown);
    expect(approval.kind).toBe("validation-warning-approval");
    expect(checkDataValidationBatch(workbook, targets, approval)).toBe(1);
  });

  test("an approval is single use", () => {
    const workbook = sampleWorkbook();
    const sheetId = workbook.sheets[0].id;
    workbook.sheets[0].validations = normalizeDataValidationRules([
      numberRule({
        id: "warn",
        range: "B1:B1",
        operator: "lessEqual",
        min: 5,
        max: 0,
        allowBlank: false,
        message: "Above 5",
        errorStyle: "warning",
      }),
    ]);
    const targets = [{ sheetId, address: "B1" }];
    let thrown;
    try {
      checkDataValidationBatch(workbook, targets);
    } catch (error) {
      thrown = error;
    }
    const approval = assertNoStopValidationIssues(thrown);
    checkDataValidationBatch(workbook, targets, approval);
    expect(() => checkDataValidationBatch(workbook, targets, approval)).toThrow(
      "Validation approval is stale or invalid. Review the complete batch again.",
    );
  });

  test("assertNoStopValidationIssues refuses a stop violation and an empty error", () => {
    const workbook = sampleWorkbook();
    const sheetId = workbook.sheets[0].id;
    workbook.sheets[0].validations = normalizeDataValidationRules([
      numberRule(),
    ]);
    let thrown;
    try {
      checkDataValidationBatch(workbook, [{ sheetId, address: "A2" }]);
    } catch (error) {
      thrown = error;
    }
    expect(() => assertNoStopValidationIssues(thrown)).toThrow(
      "Stop errors cannot be overridden. Correct the invalid values.",
    );
    expect(() =>
      assertNoStopValidationIssues(new DataValidationError(workbook, [])),
    ).toThrow("Stop errors cannot be overridden. Correct the invalid values.");
  });

  test("DataValidationError lists at most five issues in its message", () => {
    const workbook = sampleWorkbook();
    const issues = Array.from({ length: 6 }, (_, index) => ({
      sheetName: "Sheet1",
      address: `A${index + 1}`,
      message: "Enter 1 to 10",
      reason: "Value does not satisfy the rule.",
      style: "warning",
      evaluationError: false,
    }));
    const error = new DataValidationError(workbook, issues);
    expect(error.name).toBe("DataValidationError");
    expect(error.issues).toHaveLength(6);
    expect(error.message.split("\n")).toHaveLength(6);
    expect(error.message).toContain("6 validation violations in this batch.");
  });

  test("summarizeValidationInspection counts checks and invalid cells", () => {
    const workbook = sampleWorkbook();
    const sheetId = workbook.sheets[0].id;
    workbook.sheets[0].validations = normalizeDataValidationRules([
      numberRule({ range: "A1:A3" }),
    ]);
    const report = summarizeValidationInspection(workbook);
    expect(report).toMatchObject({
      checks: 3,
      invalidCells: 1,
      violations: 1,
      complete: true,
      detailsTruncated: false,
    });
    expect(report.issues[0]).toMatchObject({ sheetId, address: "A2" });
  });

  test("summarizeValidationInspection limits itself to one sheet when asked", () => {
    const workbook = sampleWorkbook();
    const sheetId = workbook.sheets[0].id;
    workbook.sheets[0].validations = normalizeDataValidationRules([
      numberRule(),
    ]);
    expect(summarizeValidationInspection(workbook, sheetId).checks).toBe(3);
    expect(
      summarizeValidationInspection(workbook, "no-such-sheet").checks,
    ).toBe(0);
  });

  test("resolveValidationInspection serialises a current report", () => {
    const workbook = sampleWorkbook();
    workbook.sheets[0].validations = normalizeDataValidationRules([
      numberRule(),
    ]);
    const report = summarizeValidationInspection(workbook);
    const text = resolveValidationInspection(workbook, report);
    expect(typeof text).toBe("string");
    expect(JSON.parse(text)).toMatchObject({
      schemaVersion: 1,
      kind: "local-data-validation-inspection",
      report: { checks: 3, invalidCells: 1 },
    });
  });

  test("resolveValidationInspection refuses a report for a changed workbook", () => {
    const workbook = sampleWorkbook();
    workbook.sheets[0].validations = normalizeDataValidationRules([
      numberRule(),
    ]);
    const report = summarizeValidationInspection(workbook);
    const changed = structuredClone(workbook);
    changed.sheets[0].cells.A9 = { raw: 1 };
    expect(() => resolveValidationInspection(changed, report)).toThrow(
      "Validation inspection is stale or unknown. Inspect existing values again.",
    );
  });
});

describe("conditional formatting rules", () => {
  test("normalizeConditionalRules returns normalised copies with upper-case ranges", () => {
    const [rule] = normalizeConditionalRules([
      conditionalRule({ range: "b1:a2" }),
    ]);
    expect(rule).toEqual(conditionalRule({ range: "A1:B2" }));
  });

  test("normalizeConditionalRules rejects a colour that is not six-digit hex", () => {
    expect(() =>
      normalizeConditionalRules([conditionalRule({ fill: "red" })]),
    ).toThrow("Choose a six-digit hexadecimal color.");
  });

  test("normalizeConditionalRules rejects repeated priorities and stop-if-true on a data bar", () => {
    expect(() =>
      normalizeConditionalRules([
        conditionalRule({ id: "a", priority: 1 }),
        conditionalRule({ id: "b", priority: 1 }),
      ]),
    ).toThrow("Conditional priorities must be unique integers from 1 to 100.");
    expect(() =>
      normalizeConditionalRules([
        conditionalRule({ kind: "dataBar", stopIfTrue: true }),
      ]),
    ).toThrow(
      "Stop if true is not supported for color scales, data bars, or icon sets.",
    );
  });

  test("normalizeConditionalRules needs a formula beginning with = for formula rules", () => {
    expect(() =>
      normalizeConditionalRules([
        conditionalRule({ kind: "formula", formula: "A1>1" }),
      ]),
    ).toThrow("A formula rule requires a formula beginning with =.");
    expect(
      normalizeConditionalRules([
        conditionalRule({ kind: "formula", formula: "=A1>1" }),
      ])[0].formula,
    ).toBe("=A1>1");
  });

  test("normalizeConditionalRules rejects decreasing icon thresholds", () => {
    expect(() =>
      normalizeConditionalRules([
        conditionalRule({ kind: "iconSet", iconThresholds: [66, 33] }),
      ]),
    ).toThrow(
      "Choose two increasing icon thresholds; percentages must be between 0 and 100.",
    );
  });

  test("conditionalRuleKinds lists the fifteen kinds, and each one normalises", () => {
    expect(conditionalRuleKinds).toHaveLength(15);
    expect(conditionalRuleKinds).toContain("iconSet");
    for (const kind of conditionalRuleKinds) {
      expect(
        normalizeConditionalRules([
          conditionalRule({ kind, formula: "=A1>0" }),
        ])[0].kind,
      ).toBe(kind);
    }
  });

  test("sortConditionalRulesByPriority puts prioritised rules first and keeps the input", () => {
    const rules = normalizeConditionalRules([
      conditionalRule({ id: "a" }),
      conditionalRule({ id: "b", priority: 1 }),
      conditionalRule({ id: "c" }),
      conditionalRule({ id: "d", priority: 2 }),
    ]);
    const sorted = sortConditionalRulesByPriority(rules);
    expect(sorted.map((rule) => rule.id)).toEqual(["b", "d", "c", "a"]);
    expect(rules.map((rule) => rule.id)).toEqual(["a", "b", "c", "d"]);
  });

  test("sortConditionalRulesByPriority lists unprioritised rules last listed first", () => {
    const rules = normalizeConditionalRules([
      conditionalRule({ id: "x" }),
      conditionalRule({ id: "y" }),
      conditionalRule({ id: "z" }),
    ]);
    expect(
      sortConditionalRulesByPriority(rules).map((rule) => rule.id),
    ).toEqual(["z", "y", "x"]);
  });

  test("normalizeConditionalRules refuses a kind that is not in conditionalRuleKinds", () => {
    expect(conditionalRuleKinds).not.toContain("bogus");
    expect(() =>
      normalizeConditionalRules([conditionalRule({ kind: "bogus" })]),
    ).toThrow("Invalid conditional rule.");
  });

  test("evaluateConditionalFormatRules styles only matching cells", () => {
    const workbook = sampleWorkbook();
    const sheet = {
      ...workbook.sheets[0],
      conditionalFormats: normalizeConditionalRules([conditionalRule()]),
    };
    const { calculator } = getWorkbookCalculation(workbook);
    const evaluate = evaluateConditionalFormatRules(sheet, calculator);
    expect(evaluate("A2").style).toEqual({ fill: "#ffcccc", color: "#000000" });
    expect(evaluate("A1").style).toEqual({});
  });

  test("evaluateConditionalFormatRules styles duplicate values and formula matches", () => {
    const workbook = sampleWorkbook();
    workbook.sheets[0].cells.A1 = { raw: 7 };
    workbook.sheets[0].cells.A3 = { raw: 7 };
    const sheet = {
      ...workbook.sheets[0],
      conditionalFormats: normalizeConditionalRules([
        conditionalRule({
          id: "dup",
          kind: "duplicate",
          range: "A1:A3",
          value: "",
          second: "",
          fill: "#ffff00",
        }),
        conditionalRule({
          id: "fx",
          kind: "formula",
          range: "B1:B1",
          formula: "=B1>1",
          value: "",
          second: "",
          fill: "#00ff00",
        }),
      ]),
    };
    const { calculator } = getWorkbookCalculation(workbook);
    const evaluate = evaluateConditionalFormatRules(sheet, calculator);
    expect(evaluate("A1").style.fill).toBe("#ffff00");
    expect(evaluate("A3").style.fill).toBe("#ffff00");
    expect(evaluate("A2").style).toEqual({});
    expect(evaluate("B1").style.fill).toBe("#00ff00");
  });

  test("evaluateConditionalFormatRules reports rules with bad operands through errors()", () => {
    const workbook = sampleWorkbook();
    const sheet = {
      ...workbook.sheets[0],
      conditionalFormats: normalizeConditionalRules([
        conditionalRule({ value: "abc" }),
      ]),
    };
    const { calculator } = getWorkbookCalculation(workbook);
    const evaluate = evaluateConditionalFormatRules(sheet, calculator);
    expect(evaluate("A1").errors).toEqual([
      {
        ruleId: "r1",
        address: "A1",
        message:
          "Numeric rule operands must be finite and minimum cannot exceed maximum.",
      },
    ]);
    expect(evaluate.errors()).toHaveLength(1);
  });

  test("updateConditionalFormatRules previews a change without altering the workbook", () => {
    const workbook = sampleWorkbook();
    workbook.sheets[0].conditionalFormats = normalizeConditionalRules([
      conditionalRule(),
    ]);
    const preview = updateConditionalFormatRules(
      workbook,
      workbook.sheets[0].id,
      [conditionalRule({ value: "20" })],
    );
    expect(preview.changes).toHaveLength(1);
    expect(preview.changes[0].before.value).toBe("10");
    expect(preview.changes[0].after.value).toBe("20");
    expect(workbook.sheets[0].conditionalFormats[0].value).toBe("10");
  });

  test("updateConditionalFormatRules refuses protected sheets and read-only workbooks", () => {
    const workbook = sampleWorkbook();
    const sheetId = workbook.sheets[0].id;
    const protectedBook = {
      ...workbook,
      sheets: [{ ...workbook.sheets[0], protection: { enabled: true } }],
    };
    expect(() =>
      updateConditionalFormatRules(protectedBook, sheetId, [conditionalRule()]),
    ).toThrow("Unprotect the worksheet before changing conditional formats.");
    expect(() =>
      updateConditionalFormatRules(workbook, sheetId, [conditionalRule()], {
        readOnly: true,
      }),
    ).toThrow("This workbook is read-only.");
  });

  test("updateConditionalFormatRules reports no changes when the rules are the same", () => {
    const workbook = sampleWorkbook();
    workbook.sheets[0].conditionalFormats = normalizeConditionalRules([
      conditionalRule(),
    ]);
    const preview = updateConditionalFormatRules(
      workbook,
      workbook.sheets[0].id,
      [conditionalRule()],
    );
    expect(preview.changes).toEqual([]);
  });

  test("assertConditionalFormatPreviewCurrent applies a current preview to a new workbook", () => {
    const workbook = sampleWorkbook();
    workbook.sheets[0].conditionalFormats = normalizeConditionalRules([
      conditionalRule(),
    ]);
    const preview = updateConditionalFormatRules(
      workbook,
      workbook.sheets[0].id,
      [conditionalRule({ value: "20" })],
    );
    const next = assertConditionalFormatPreviewCurrent(workbook, preview);
    expect(next.sheets[0].conditionalFormats[0].value).toBe("20");
    expect(workbook.sheets[0].conditionalFormats[0].value).toBe("10");
    expect(next.revision).toBe(workbook.revision);
  });

  test("assertConditionalFormatPreviewCurrent refuses a preview made from an older workbook", () => {
    const workbook = sampleWorkbook();
    workbook.sheets[0].conditionalFormats = normalizeConditionalRules([
      conditionalRule(),
    ]);
    const preview = updateConditionalFormatRules(
      workbook,
      workbook.sheets[0].id,
      [conditionalRule({ value: "20" })],
    );
    const changed = structuredClone(workbook);
    changed.sheets[0].cells.C1 = { raw: 1 };
    expect(() =>
      assertConditionalFormatPreviewCurrent(changed, preview),
    ).toThrow(
      "Conditional format preview is stale. Reopen the manager to refresh it.",
    );
  });
});

describe("defined names", () => {
  function workbookWithRate() {
    const workbook = sampleWorkbook();
    const preview = applyDefinedNameOperation(workbook, {
      kind: "create",
      definition: { name: "Rate", refersTo: "=Sheet1!$A$1" },
    });
    return { workbook, next: assertNamePreviewCurrent(workbook, preview) };
  }

  test("applyDefinedNameOperation previews a create without changing the workbook", () => {
    const workbook = sampleWorkbook();
    const preview = applyDefinedNameOperation(workbook, {
      kind: "create",
      definition: { name: "Rate", refersTo: "=Sheet1!$A$1" },
    });
    expect(preview.after).toEqual({ name: "Rate", refersTo: "=Sheet1!$A$1" });
    expect(preview.changes).toEqual([]);
    expect(workbook.definedNames).toBeUndefined();
  });

  test("applyDefinedNameOperation refuses a name that already exists in its scope", () => {
    const { next } = workbookWithRate();
    expect(() =>
      applyDefinedNameOperation(next, {
        kind: "create",
        definition: { name: "rate", refersTo: "=1" },
      }),
    ).toThrow("A name already exists in this scope.");
  });

  test("applyDefinedNameOperation refuses invalid names and read-only workbooks", () => {
    const workbook = sampleWorkbook();
    expect(() =>
      applyDefinedNameOperation(workbook, {
        kind: "create",
        definition: { name: "A1", refersTo: "=1" },
      }),
    ).toThrow(
      "Names must begin with a letter or underscore and cannot be cell addresses.",
    );
    expect(() =>
      applyDefinedNameOperation(
        workbook,
        { kind: "create", definition: { name: "Rate", refersTo: "=1" } },
        { readOnly: true },
      ),
    ).toThrow("This workbook is read-only.");
  });

  test("a rename rewrites the formulas that use the name", () => {
    const { next } = workbookWithRate();
    next.sheets[0].cells.B1 = { raw: "=Rate*2" };
    const preview = applyDefinedNameOperation(next, {
      kind: "rename",
      name: "Rate",
      newName: "Price",
    });
    expect(preview.after).toEqual({ name: "Price", refersTo: "=Sheet1!$A$1" });
    expect(preview.changes).toEqual([
      expect.objectContaining({
        address: "B1",
        before: "=Rate*2",
        after: "=Price*2",
      }),
    ]);
    expect(preview.next.sheets[0].cells.B1).toEqual({ raw: "=Price*2" });
  });

  test("a delete is refused while a formula uses the name", () => {
    const { next } = workbookWithRate();
    next.sheets[0].cells.B1 = { raw: "=Rate*2" };
    expect(() =>
      applyDefinedNameOperation(next, { kind: "delete", name: "Rate" }),
    ).toThrow(
      "This name is referenced by Sheet1!B1. Remove or replace its dependencies before deletion.",
    );
  });

  test("an update changes the definition and the comment", () => {
    const { next } = workbookWithRate();
    const preview = applyDefinedNameOperation(next, {
      kind: "update",
      name: "Rate",
      definition: {
        name: "Rate",
        refersTo: "=Sheet1!$A$1*10",
        comment: "tenfold",
      },
    });
    expect(preview.before).toEqual({ name: "Rate", refersTo: "=Sheet1!$A$1" });
    expect(preview.after).toEqual({
      name: "Rate",
      refersTo: "=Sheet1!$A$1*10",
      comment: "tenfold",
    });
  });

  test("assertNamePreviewCurrent returns the workbook with the name added", () => {
    const workbook = sampleWorkbook();
    const preview = applyDefinedNameOperation(workbook, {
      kind: "create",
      definition: { name: "Rate", refersTo: "=Sheet1!$A$1" },
    });
    const next = assertNamePreviewCurrent(workbook, preview);
    expect(next.definedNames).toEqual([
      { name: "Rate", refersTo: "=Sheet1!$A$1" },
    ]);
    expect(workbook.definedNames).toBeUndefined();
  });

  test("assertNamePreviewCurrent refuses a preview whose workbook has changed", () => {
    const workbook = sampleWorkbook();
    const preview = applyDefinedNameOperation(workbook, {
      kind: "create",
      definition: { name: "Rate", refersTo: "=Sheet1!$A$1" },
    });
    const changed = structuredClone(workbook);
    changed.sheets[0].cells.C1 = { raw: 2 };
    expect(() => assertNamePreviewCurrent(changed, preview)).toThrow(
      "The workbook changed after the name preview. Review the operation again.",
    );
  });

  test("resolveDefinedNameValue expands a name and leaves unknown names alone", () => {
    const { next } = workbookWithRate();
    const sheetId = next.sheets[0].id;
    expect(resolveDefinedNameValue("=Rate+1", next, sheetId)).toBe(
      "=(Sheet1!$A$1)+1",
    );
    expect(resolveDefinedNameValue("=Nope*2", next, sheetId)).toBe("=Nope*2");
  });

  test("resolveDefinedNameValue reports a circular name and a refused sheet", () => {
    const { next } = workbookWithRate();
    const sheetId = next.sheets[0].id;
    expect(() =>
      resolveDefinedNameValue("=Rate+1", next, sheetId, [":rate"]),
    ).toThrow("#CYCLE!");
    expect(() =>
      resolveDefinedNameValue("=Rate+1", next, sheetId, [], () => false),
    ).toThrow("#BLOCKED!");
  });
});

describe("review threads, hyperlinks and formula watches", () => {
  /** A workbook with one open review thread on A1 that has one message. */
  function workbookWithThread() {
    const workbook = sampleWorkbook();
    const sheet = workbook.sheets[0];
    const thread = {
      id: "t1",
      anchor: { sheetId: sheet.id, sheetName: sheet.name, address: "A1" },
      createdAt: "2026-01-01T10:00:00.000Z",
      updatedAt: "2026-01-01T10:00:00.000Z",
      messages: [
        {
          id: "m1",
          author: "Ann",
          text: "Check this value",
          createdAt: "2026-01-01T10:00:00.000Z",
        },
      ],
    };
    return { ...workbook, reviewThreads: [thread] };
  }

  test("addReviewThreadReply appends a message and leaves the original unchanged", () => {
    const workbook = workbookWithThread();
    const next = addReviewThreadReply(workbook, "t1", "Looks fine", "Ben", {
      timestamp: STAMP,
    });
    expect(
      next.reviewThreads[0].messages.map((message) => message.author),
    ).toEqual(["Ann", "Ben"]);
    expect(next.reviewThreads[0].updatedAt).toBe(STAMP);
    expect(workbook.reviewThreads[0].messages).toHaveLength(1);
  });

  test("addReviewThreadReply uses Reviewer as the default author", () => {
    const next = addReviewThreadReply(
      workbookWithThread(),
      "t1",
      "Second",
      undefined,
      {
        timestamp: STAMP,
      },
    );
    expect(next.reviewThreads[0].messages.at(-1).author).toBe("Reviewer");
  });

  test("addReviewThreadReply refuses empty, overlong, unknown and resolved threads", () => {
    const workbook = workbookWithThread();
    expect(() => addReviewThreadReply(workbook, "t1", "   ", "Ben")).toThrow(
      "Reply must contain 1–4000 characters without unsupported control characters.",
    );
    expect(() =>
      addReviewThreadReply(
        workbook,
        "t1",
        "x".repeat(reviewThreadLimits.textLength + 1),
        "Ben",
      ),
    ).toThrow("Reply must contain 1–4000 characters");
    expect(() =>
      addReviewThreadReply(workbook, "missing", "Hi", "Ben"),
    ).toThrow("The discussion could not be found. Reopen the review tools.");
    const resolved = setReviewThreadResolved(workbook, "t1", true, "Cy", {
      timestamp: STAMP,
    });
    expect(() => addReviewThreadReply(resolved, "t1", "Hi", "Ben")).toThrow(
      "Reopen the discussion before adding a reply.",
    );
  });

  test("setReviewThreadResolved resolves and then reopens a thread", () => {
    const workbook = workbookWithThread();
    const resolved = setReviewThreadResolved(workbook, "t1", true, "Cy", {
      timestamp: STAMP,
    });
    expect(resolved.reviewThreads[0].resolved).toEqual({
      author: "Cy",
      createdAt: STAMP,
    });
    expect(workbook.reviewThreads[0].resolved).toBeUndefined();
    const reopened = setReviewThreadResolved(resolved, "t1", false, "Cy", {
      timestamp: "2026-01-03T00:00:00.000Z",
    });
    expect("resolved" in reopened.reviewThreads[0]).toBe(false);
  });

  test("setReviewThreadResolved returns the same workbook when nothing changes", () => {
    const workbook = workbookWithThread();
    expect(setReviewThreadResolved(workbook, "t1", false)).toBe(workbook);
    expect(() => setReviewThreadResolved(workbook, "t1", "yes")).toThrow(
      "Choose whether to resolve or reopen the discussion.",
    );
  });

  test("addReviewThreadReply accepts a reply of exactly textLength characters", () => {
    const workbook = workbookWithThread();
    const text = "x".repeat(reviewThreadLimits.textLength);
    const next = addReviewThreadReply(workbook, "t1", text, "Ben", {
      timestamp: STAMP,
    });
    expect(next.reviewThreads[0].messages.at(-1).text).toBe(text);
    expect(() =>
      addReviewThreadReply(workbook, "t1", `${text}x`, "Ben", {
        timestamp: STAMP,
      }),
    ).toThrow("Reply must contain 1–4000 characters");
  });

  test("reviewThreadLimits gives the published limits", () => {
    expect(reviewThreadLimits).toEqual({
      threads: 1000,
      messagesPerThread: 100,
      totalMessages: 5000,
      textLength: 4000,
      authorLength: 100,
      urlLength: 2048,
      labelLength: 256,
    });
  });

  test("validateHyperlinkRecord normalises a web link and keeps its label", () => {
    expect(
      validateHyperlinkRecord({
        kind: "external",
        url: "https://example.com/a b?x=1",
        label: "Example",
      }),
    ).toEqual({
      kind: "external",
      url: "https://example.com/a%20b?x=1",
      label: "Example",
    });
  });

  test("validateHyperlinkRecord accepts a plain mailto address and refuses other schemes", () => {
    expect(
      validateHyperlinkRecord({
        kind: "external",
        url: "mailto:a@example.com",
      }),
    ).toEqual({
      kind: "external",
      url: "mailto:a@example.com",
    });
    expect(() =>
      validateHyperlinkRecord({ kind: "external", url: "javascript:alert(1)" }),
    ).toThrow("Only https://, http:// and mailto: links are supported.");
    expect(() =>
      validateHyperlinkRecord({
        kind: "external",
        url: "mailto:a@example.com?subject=hi",
      }),
    ).toThrow("Use mailto: with one email address and no additional headers.");
  });

  test("validateHyperlinkRecord refuses credentials and bad worksheet addresses", () => {
    expect(() =>
      validateHyperlinkRecord({
        kind: "external",
        url: "https://user:pw@example.com",
      }),
    ).toThrow("Links cannot contain embedded credentials.");
    expect(
      validateHyperlinkRecord({
        kind: "internal",
        sheetId: "sheet-1",
        address: "b2",
        label: "Go",
      }),
    ).toEqual({
      kind: "internal",
      sheetId: "sheet-1",
      address: "B2",
      label: "Go",
    });
  });

  test("setCellHyperlink sets a link on a cell and keeps the value", () => {
    const workbook = sampleWorkbook();
    const sheetId = workbook.sheets[0].id;
    workbook.sheets[0].cells.A3 = { raw: 7 };
    const next = setCellHyperlink(workbook, sheetId, "A3", {
      kind: "external",
      url: "https://example.com",
    });
    expect(next.sheets[0].cells.A3).toEqual({
      raw: 7,
      hyperlink: { kind: "external", url: "https://example.com/" },
    });
    expect(workbook.sheets[0].cells.A3).toEqual({ raw: 7 });
  });

  test("setCellHyperlink leaves the revision unchanged and refuses bad input", () => {
    const workbook = { ...sampleWorkbook(), revision: 3 };
    const sheetId = workbook.sheets[0].id;
    const next = setCellHyperlink(workbook, sheetId, "A3", {
      kind: "external",
      url: "https://example.com",
    });
    expect(next.revision).toBe(3);
    expect(() =>
      setCellHyperlink(workbook, "no-such-sheet", "A3", {
        kind: "external",
        url: "https://example.com",
      }),
    ).toThrow("Sheet not found: no-such-sheet");
    expect(() =>
      setCellHyperlink(
        workbook,
        sheetId,
        "A3",
        {
          kind: "external",
          url: "https://example.com",
        },
        { readOnly: true },
      ),
    ).toThrow("This workbook is read-only.");
  });

  test("validateFormulaWatches returns watches, and marks one with a missing sheet as orphaned", () => {
    const workbook = sampleWorkbook();
    const sheet = workbook.sheets[0];
    expect(validateFormulaWatches(undefined, workbook)).toEqual([]);
    expect(
      validateFormulaWatches(
        [
          {
            id: "w1",
            anchor: { sheetId: sheet.id, sheetName: sheet.name, address: "a1" },
          },
        ],
        workbook,
      ),
    ).toEqual([
      {
        id: "w1",
        anchor: { sheetId: sheet.id, sheetName: sheet.name, address: "A1" },
      },
    ]);
    expect(
      validateFormulaWatches(
        [
          {
            id: "w2",
            anchor: { sheetId: "gone", sheetName: "Old", address: "C3" },
          },
        ],
        workbook,
      ),
    ).toEqual([
      {
        id: "w2",
        anchor: { sheetId: "gone", sheetName: "Old", address: "C3" },
        orphaned: "missing-sheet",
      },
    ]);
  });

  test("validateFormulaWatches refuses two active watches on one cell and repeated IDs", () => {
    const workbook = sampleWorkbook();
    const sheet = workbook.sheets[0];
    const watch = {
      id: "w1",
      anchor: { sheetId: sheet.id, sheetName: sheet.name, address: "A1" },
    };
    expect(() =>
      validateFormulaWatches(
        [
          watch,
          { ...watch, id: "w3", anchor: { ...watch.anchor, address: "a1" } },
        ],
        workbook,
      ),
    ).toThrow("Active formula watches must not repeat the same cell.");
    expect(() =>
      validateFormulaWatches([watch, { ...watch }], workbook),
    ).toThrow("Formula watch IDs must be unique.");
  });
});
