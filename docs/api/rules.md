# Rules: validation, conditional formats, names and review

This area covers the rules that sit on a worksheet or workbook: data validation, conditional formatting, defined names, review threads, hyperlinks and formula watches. Most functions take a workbook and either return a new workbook or a preview of a change. None of them change the workbook they are given. Functions that check values or formulas do so without storing anything.

## normalizeDataValidationRules

```ts
export function normalizeDataValidationRules(
  validations: DataValidationRule[],
): DataValidationRule[];
```

Returns normalised copies of the rules, with ranges in upper case and unknown fields dropped. It throws if a rule is invalid, if IDs repeat, or if there are more than 100 rules. Store the result as `sheet.validations`.

```js
import {
  createBlankWorkbook,
  normalizeDataValidationRules,
} from "@ledgerglass/workbook";

const workbook = createBlankWorkbook("Budget");
workbook.sheets[0].validations = normalizeDataValidationRules([
  {
    id: "amount",
    range: "C2:C50",
    kind: "number",
    operator: "greaterEqual",
    min: 0,
    max: 0,
    list: [],
    allowBlank: true,
    message: "Enter an amount of zero or more.",
  },
]);
```

## validateDataValidationRule

```ts
export function validateDataValidationRule(
  rule: DataValidationRule,
  value: unknown,
  datePolicy?: DatePolicy,
): boolean;
```

Returns true if the value satisfies the rule and false if it does not. An empty value is allowed only when the rule has `allowBlank: true`. Custom rules always return false here, because they need the workbook; check them with `checkDataValidationBatch`.

```js
import {
  normalizeDataValidationRules,
  validateDataValidationRule,
} from "@ledgerglass/workbook";

const [rule] = normalizeDataValidationRules([
  {
    id: "score",
    range: "A1:A10",
    kind: "number",
    operator: "between",
    min: 1,
    max: 10,
    list: [],
    allowBlank: false,
    message: "Enter 1 to 10.",
  },
]);
validateDataValidationRule(rule, 7); // true
validateDataValidationRule(rule, 11); // false
```

## findDataValidationsInRange

```ts
export function findDataValidationsInRange(
  workbook: Workbook,
  sheetId: string,
  cellAddress: string,
): DataValidationRule[];
```

Returns the validation rules on a worksheet whose range covers the given cell. It returns an empty array when no rule covers the cell, and it throws if the address is not valid. The cell address is case-insensitive.

## checkDataValidationBatch

```ts
export function checkDataValidationBatch(
  book: Workbook,
  targets: ValidationTarget[],
  validationApproval?: ValidationApproval,
): number;
```

Checks each target cell against the validation rules that cover it. It returns 0 when every value passes, and throws a `DataValidationError` when a check fails and no approval covers the failure. When every failure is a warning, an approval from `assertNoStopValidationIssues` lets the batch go through, and the function then returns the number of issues approved.

```js
import {
  assertNoStopValidationIssues,
  checkDataValidationBatch,
  createBlankWorkbook,
  normalizeDataValidationRules,
} from "@ledgerglass/workbook";

const workbook = createBlankWorkbook("Budget");
const sheet = workbook.sheets[0];
sheet.cells.B2 = { raw: 12 };
sheet.validations = normalizeDataValidationRules([
  {
    id: "rate",
    range: "B2:B2",
    kind: "number",
    operator: "lessEqual",
    min: 10,
    max: 0,
    list: [],
    allowBlank: false,
    message: "Rates above 10 need approval.",
    errorStyle: "warning",
  },
]);
const targets = [{ sheetId: sheet.id, address: "B2" }];

try {
  checkDataValidationBatch(workbook, targets);
} catch (error) {
  // Show error.issues to the user, then approve the warnings.
  const approval = assertNoStopValidationIssues(error);
  checkDataValidationBatch(workbook, targets, approval); // returns 1
}
```

## assertNoStopValidationIssues

```ts
export function assertNoStopValidationIssues(
  error: DataValidationErrorInstance,
): ValidationApproval;
```

Turns a `DataValidationError` into an approval for `checkDataValidationBatch`. It throws if the error has no issues, or if any issue has the `stop` style or could not be evaluated. An approval works once; a second use throws.

## DataValidationError

```ts
export const DataValidationError: new (
  workbook: Workbook,
  issues: DataValidationIssue[],
) => DataValidationErrorInstance;
```

The error thrown by `checkDataValidationBatch` when a check fails. Its `issues` array lists each failed check, and its message gives the first five. Hosts receive it from the check and do not usually construct it.

## summarizeValidationInspection

```ts
export function summarizeValidationInspection(
  workbook: Workbook,
  sheetId?: string,
): ValidationInspectionReport;
```

Checks every value on the workbook, or on one sheet when a sheet ID is given, and returns a report with the number of checks, invalid cells and violations. The report includes the failed issues. The workbook is not changed.

## resolveValidationInspection

```ts
export function resolveValidationInspection(
  workbook: Workbook,
  report: ValidationInspectionReport,
): string;
```

Returns the report as JSON text with `schemaVersion` and `kind` fields. It throws if the workbook has changed since the report was made, so the report must be regenerated.

## normalizeConditionalRules

```ts
export function normalizeConditionalRules(
  rules: ConditionalRule[],
): ConditionalRule[];
```

Returns normalised copies of conditional formatting rules, with ranges in upper case. It throws if a rule is invalid, if an ID or priority repeats, or if a colour is not a six-digit hex value. The result is what you store as `sheet.conditionalFormats`.

```js
import {
  createBlankWorkbook,
  normalizeConditionalRules,
} from "@ledgerglass/workbook";

const workbook = createBlankWorkbook("Budget");
workbook.sheets[0].conditionalFormats = normalizeConditionalRules([
  {
    id: "over-budget",
    range: "C2:C50",
    kind: "greater",
    value: "1000",
    second: "",
    fill: "#ffcccc",
    color: "#7a0000",
    priority: 1,
  },
]);
```

## evaluateConditionalFormatRules

```ts
export function evaluateConditionalFormatRules(
  sheet: Sheet,
  calculator: ConditionalCalculator,
  cells?: Record<string, Cell>,
): ConditionalFormatEvaluator;
```

Returns a function that gives the style for one cell address, and a `errors()` method that lists rule problems found so far. The calculator must be the callable from `getWorkbookCalculation`. Formula rules and rules on blank cells are both evaluated.

```js
import {
  createBlankWorkbook,
  evaluateConditionalFormatRules,
  getWorkbookCalculation,
  normalizeConditionalRules,
} from "@ledgerglass/workbook";

const workbook = createBlankWorkbook("Budget");
const sheet = workbook.sheets[0];
sheet.cells.C2 = { raw: 1500 };
sheet.conditionalFormats = normalizeConditionalRules([
  {
    id: "over-budget",
    range: "C2:C50",
    kind: "greater",
    value: "1000",
    second: "",
    fill: "#ffcccc",
    color: "#7a0000",
  },
]);
const { calculator } = getWorkbookCalculation(workbook);
const evaluate = evaluateConditionalFormatRules(sheet, calculator);
evaluate("C2").style; // { fill: "#ffcccc", color: "#7a0000" }
```

## updateConditionalFormatRules

```ts
export function updateConditionalFormatRules(
  workbook: Workbook,
  sheetId: string,
  rules: ConditionalRule[],
  options?: RuleEditOptions,
): ConditionalFormatPreview;
```

Returns a preview of replacing a worksheet's conditional formats with the given rules. The preview lists the changed rules and sample cells with their before and after styles. The workbook is not changed; apply the preview with `assertConditionalFormatPreviewCurrent`. It throws if the sheet is missing or protected, or if the workbook is read-only.

## assertConditionalFormatPreviewCurrent

```ts
export function assertConditionalFormatPreviewCurrent(
  workbook: Workbook,
  preview: ConditionalFormatPreview,
  options?: RuleEditOptions,
): Workbook;
```

Returns a new workbook with the previewed rules applied, or the same workbook when the preview has no changes. In the first case the `revision` is one higher and `events` has one new event, `Apply conditional format change`, with an empty change list. In the second nothing is recorded. It throws if the workbook has changed since the preview was made.

```js
import {
  assertConditionalFormatPreviewCurrent,
  createBlankWorkbook,
  normalizeConditionalRules,
  updateConditionalFormatRules,
} from "@ledgerglass/workbook";

const workbook = createBlankWorkbook("Budget");
const sheetId = workbook.sheets[0].id;
const preview = updateConditionalFormatRules(workbook, sheetId, [
  {
    id: "over-budget",
    range: "C2:C50",
    kind: "greater",
    value: "2000",
    second: "",
    fill: "#ffcccc",
    color: "#7a0000",
  },
]);
const next = assertConditionalFormatPreviewCurrent(workbook, preview);
```

## sortConditionalRulesByPriority

```ts
export function sortConditionalRulesByPriority(
  rules: ConditionalRule[],
): ConditionalRule[];
```

Returns a sorted copy of the rules. Rules with a `priority` come first in ascending order; rules without one follow, last listed first. The input array is not changed.

## conditionalRuleKinds

```ts
export const conditionalRuleKinds: readonly ConditionalRuleKind[];
```

The fifteen conditional rule kinds the model supports, such as `"greater"`, `"duplicate"`, `"formula"` and `"iconSet"`. Use it to build a list of kinds in an editor. It is not changed by the functions in this area.

## applyDefinedNameOperation

```ts
export function applyDefinedNameOperation(
  workbook: Workbook,
  operation: DefinedNameOperation,
  options?: RuleEditOptions,
): DefinedNamePreview;
```

Previews a create, update, rename or delete of a defined name. The preview lists the formulas that a rename would rewrite and the new workbook as `next`. It throws if the name is invalid or already exists, if a delete would leave a formula broken, or if the workbook is read-only. Apply the preview with `assertNamePreviewCurrent`.

```js
import {
  applyDefinedNameOperation,
  assertNamePreviewCurrent,
  createBlankWorkbook,
} from "@ledgerglass/workbook";

const workbook = createBlankWorkbook("Budget");
const sheet = workbook.sheets[0];
sheet.cells.A1 = { raw: 0.2 };
sheet.cells.B1 = { raw: "=Rate*100" };

const created = applyDefinedNameOperation(workbook, {
  kind: "create",
  definition: { name: "Rate", refersTo: "=Sheet1!$A$1" },
});
const withRate = assertNamePreviewCurrent(workbook, created);

const renamed = applyDefinedNameOperation(withRate, {
  kind: "rename",
  name: "Rate",
  newName: "TaxRate",
});
renamed.changes; // the B1 formula changes from "=Rate*100" to "=TaxRate*100"
```

## assertNamePreviewCurrent

```ts
export function assertNamePreviewCurrent(
  workbook: Workbook,
  preview: DefinedNamePreview,
  options?: RuleEditOptions,
): Workbook;
```

Returns the workbook with a defined-name preview applied. It throws if the workbook has changed since the preview was made, or if the preview no longer matches what the operation would produce now. The function does not change its argument.

## resolveDefinedNameValue

```ts
export function resolveDefinedNameValue(
  formulaText: string,
  workbook: Workbook,
  currentSheetId: string,
  activeNameKeys?: string[],
  canReadSheet?: (sheetId: string) => boolean,
): string;
```

Returns the formula with each defined name replaced by its definition, in brackets. Names that are not defined are left as they are. It throws `#CYCLE!` for a name that refers back to itself and `#BLOCKED!` when `canReadSheet` returns false for a sheet the name points at. The optional arguments are for the recursion; a host normally passes only the first three.

## addReviewThreadReply

```ts
export function addReviewThreadReply(
  workbook: Workbook,
  reviewThreadId: string,
  replyText: string,
  authorName?: string,
  options?: ReviewEditOptions,
): Workbook;
```

Returns a new workbook with a message added to a review thread. The author defaults to `"Reviewer"`. It throws if the thread does not exist, is resolved or orphaned, or if the text is empty or longer than `reviewThreadLimits.textLength`.

```js
import {
  addReviewThreadReply,
  createBlankWorkbook,
} from "@ledgerglass/workbook";

const workbook = createBlankWorkbook("Budget");
const sheet = workbook.sheets[0];
const withThread = {
  ...workbook,
  reviewThreads: [
    {
      id: "t1",
      anchor: { sheetId: sheet.id, sheetName: sheet.name, address: "C2" },
      createdAt: "2026-01-01T10:00:00.000Z",
      updatedAt: "2026-01-01T10:00:00.000Z",
      messages: [
        {
          id: "m1",
          author: "Ann",
          text: "Confirm this figure.",
          createdAt: "2026-01-01T10:00:00.000Z",
        },
      ],
    },
  ],
};
const next = addReviewThreadReply(withThread, "t1", "Confirmed.", "Ben", {
  timestamp: "2026-01-02T09:00:00.000Z",
});
```

## setReviewThreadResolved

```ts
export function setReviewThreadResolved(
  workbook: Workbook,
  reviewThreadId: string,
  shouldResolve: boolean,
  authorName?: string,
  options?: ReviewEditOptions,
): Workbook;
```

Returns a new workbook with a review thread resolved or reopened. It returns the same workbook when the thread is already in the requested state. It throws if the thread does not exist, if `shouldResolve` is not a boolean, or if the workbook is read-only.

## reviewThreadLimits

```ts
export const reviewThreadLimits: Readonly<ReviewThreadLimits>;
```

The published limits: 1,000 threads, 100 messages per thread, 5,000 messages in all, 4,000 characters for a message, 100 characters for an author name, 2,048 characters for a URL and 256 characters for a link label.

## validateHyperlinkRecord

```ts
export function validateHyperlinkRecord(
  hyperlink: HyperlinkInput,
): HyperlinkRecord;
```

Returns a checked copy of a hyperlink. External links must be `https://`, `http://` or a single-address `mailto:`; internal links point to a sheet and address. It throws for other schemes, embedded credentials, control characters or backslashes.

```js
import { validateHyperlinkRecord } from "@ledgerglass/workbook";

validateHyperlinkRecord({
  kind: "external",
  url: "https://example.com/a b",
  label: "Example",
});
// { kind: "external", url: "https://example.com/a%20b", label: "Example" }
```

## validateFormulaWatches

```ts
export function validateFormulaWatches(
  formulaWatches: FormulaWatch[] | undefined,
  workbook: Workbook,
): FormulaWatch[];
```

Returns checked copies of formula watches. A watch whose sheet is missing is kept and marked `orphaned: "missing-sheet"`. It returns an empty array for `undefined`, and throws for repeated IDs, for two active watches on one cell, or for more than 100 watches.
