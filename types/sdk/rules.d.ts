import type { Cell, Calculator, Sheet, Workbook } from "../index";

/** Calculator that can also be called with a sheet ID and an address to read a value. */
export type ConditionalCalculator = Calculator &
  ((sheetId: string, address: string) => unknown);

/** Date settings a workbook uses to read text dates. */
export interface DatePolicy {
  calendar: string;
  dateSystem: string;
  textDateOrder: string;
  [extra: string]: unknown;
}

export type DataValidationKind =
  "list" | "number" | "integer" | "textLength" | "date" | "time" | "custom";

export type DataValidationOperator =
  | "between"
  | "notBetween"
  | "greater"
  | "greaterEqual"
  | "less"
  | "lessEqual"
  | "equal"
  | "notEqual";

/** A data validation rule. Input and normalised output have the same shape. */
export interface DataValidationRule {
  /** Unique within the worksheet, 1 to 100 characters. */
  id: string;
  /** Range such as "A1:A5"; normalised to upper case with the top-left cell first. */
  range: string;
  kind: DataValidationKind;
  operator: DataValidationOperator;
  /** Lower bound, or the value for equal and notEqual. Must be a finite number. */
  min: number;
  /** Upper bound for between and notBetween. Must be a finite number. */
  max: number;
  /** Allowed items for list rules; empty for other kinds. */
  list: string[];
  allowBlank: boolean;
  /** Error message; required, may be an empty string. */
  message: string;
  /** Formula for custom rules; starts with "=" and is relative to the top-left cell. */
  formula?: string;
  errorStyle?: "stop" | "warning" | "information";
  errorTitle?: string;
  inputTitle?: string;
  inputMessage?: string;
  showInputMessage?: boolean;
  showDropdown?: boolean;
  [extra: string]: unknown;
}

/** One failed check reported by checkDataValidationBatch or summarizeValidationInspection. */
export interface DataValidationIssue {
  sheetId: string;
  sheetName: string;
  address: string;
  ruleId: string;
  /** "stop" blocks the change; "warning" and "information" can be approved. */
  style: "stop" | "warning" | "information";
  value: unknown;
  title: string;
  message: string;
  reason: string;
  formula?: string;
  /** True when the rule could not be evaluated; such issues cannot be approved. */
  evaluationError: boolean;
  [extra: string]: unknown;
}

/** Thrown by checkDataValidationBatch when a batch has failed checks. */
export interface DataValidationErrorInstance extends Error {
  name: "DataValidationError";
  issues: DataValidationIssue[];
}

/** Opaque token that lets one batch with warnings-only issues proceed. Single use. */
export type ValidationApproval = Readonly<{
  kind: "validation-warning-approval";
}>;

/** A cell to check, given by sheet ID and address. */
export interface ValidationTarget {
  sheetId: string;
  address: string;
}

export interface ValidationInspectionReport {
  workbookId: string;
  revision: number;
  /** Number of rule-and-cell checks performed. */
  checks: number;
  invalidCells: number;
  violations: number;
  issues: DataValidationIssue[];
  /** False when the check budget or the issue list was cut short. */
  complete: boolean;
  detailsTruncated: boolean;
  message: string;
  [extra: string]: unknown;
}

export type ConditionalRuleKind =
  | "greater"
  | "less"
  | "equal"
  | "between"
  | "contains"
  | "duplicate"
  | "unique"
  | "blank"
  | "notBlank"
  | "error"
  | "notError"
  | "formula"
  | "colorScale"
  | "dataBar"
  | "iconSet";

/** A conditional formatting rule. `value` and `second` are text; use "" when unused. */
export interface ConditionalRule {
  /** Unique within the worksheet, 1 to 100 characters. */
  id: string;
  range: string;
  kind: ConditionalRuleKind;
  /** Operand for greater, less, equal, between, and contains. Empty string when unused. */
  value: string;
  /** Upper operand for between. Empty string when unused. */
  second: string;
  /** Six-digit hex fill colour, such as "#ffcc00". */
  fill: string;
  /** Six-digit hex text colour. */
  color: string;
  /** Required when kind is "formula"; starts with "=". */
  formula?: string;
  /** Unique integer from 1 to 100. Lower numbers are applied first. */
  priority?: number;
  stopIfTrue?: boolean;
  minColor?: string;
  maxColor?: string;
  thresholdType?: "percent" | "number";
  /** Two increasing thresholds for icon sets. */
  iconThresholds?: [number, number];
  [extra: string]: unknown;
}

export interface ConditionalRuleError {
  ruleId: string;
  /** Empty when the problem applies to the whole rule. */
  address: string;
  message: string;
}

/** Style that applies to one cell. */
export interface ConditionalCellStyle {
  fill?: string;
  color?: string;
  [extra: string]: unknown;
}

export interface ConditionalCellResult {
  style: ConditionalCellStyle;
  errors: ConditionalRuleError[];
}

/** Evaluates one cell when called with an address; errors() lists every problem found so far. */
export type ConditionalFormatEvaluator = ((
  address: string,
) => ConditionalCellResult) & {
  errors(): ConditionalRuleError[];
};

export interface ConditionalRuleChange {
  id: string;
  before?: ConditionalRule;
  beforeIndex?: number;
  after?: ConditionalRule;
  afterIndex?: number;
}

export interface ConditionalFormatSample {
  address: string;
  value: unknown;
  before: ConditionalCellResult;
  after: ConditionalCellResult;
}

/** Result of updateConditionalFormatRules. It describes the change and does not apply it. */
export interface ConditionalFormatPreview {
  workbookId: string;
  revision: number;
  /** Serialised workbook the preview was made from. */
  snapshotGuard: string;
  sheetId: string;
  rules: ConditionalRule[];
  changes: ConditionalRuleChange[];
  samples: ConditionalFormatSample[];
  errors: ConditionalRuleError[];
  previewLimit: number;
  [extra: string]: unknown;
}

/** Defined name entry as stored on a workbook. */
export interface DefinedNameDefinition {
  /** 1 to 255 characters; letters, digits, underscores and periods. Not a cell address. */
  name: string;
  /** Formula beginning with "=", such as "=Sheet1!$A$1". */
  refersTo: string;
  /** ID of the worksheet the name is scoped to. Omit for a workbook-wide name. */
  scopeSheetId?: string;
  /** Up to 1,000 characters. */
  comment?: string;
}

export type DefinedNameOperation =
  | { kind: "create"; definition: DefinedNameDefinition }
  | {
      kind: "update";
      name: string;
      scopeSheetId?: string;
      definition: DefinedNameDefinition;
    }
  | { kind: "rename"; name: string; scopeSheetId?: string; newName: string }
  | { kind: "delete"; name: string; scopeSheetId?: string };

export interface DefinedNameChange {
  kind: string;
  label: string;
  sheetId?: string;
  address?: string;
  before: string;
  after: string;
  [extra: string]: unknown;
}

/** Result of applyDefinedNameOperation. Use assertNamePreviewCurrent to apply it. */
export interface DefinedNamePreview {
  /** Serialised workbook the preview was made from. */
  snapshot: string;
  request: DefinedNameOperation;
  before?: DefinedNameDefinition;
  after?: DefinedNameDefinition;
  /** Formulas and definitions that would be rewritten. */
  changes: DefinedNameChange[];
  dependencies: Array<{ label: string; sheetId?: string; address?: string }>;
  /** The workbook after the operation. */
  next: Workbook;
  [extra: string]: unknown;
}

export interface ReviewMessage {
  id: string;
  author: string;
  text: string;
  createdAt: string;
}

export interface ReviewThread {
  id: string;
  anchor: { sheetId: string; sheetName: string; address: string };
  /** ISO UTC timestamp. */
  createdAt: string;
  /** ISO UTC timestamp, never earlier than createdAt. */
  updatedAt: string;
  messages: ReviewMessage[];
  /** Present when the discussion is resolved. */
  resolved?: { author: string; createdAt: string };
  /** Present when the target was deleted; such a thread cannot take replies. */
  orphaned?: string;
  [extra: string]: unknown;
}

export interface ReviewEditOptions {
  /** ISO UTC timestamp for the change. Defaults to the current time. */
  timestamp?: string;
  readOnly?: boolean;
}

export interface RuleEditOptions {
  readOnly?: boolean;
}

/** Limits enforced on review threads, messages and hyperlinks. */
export interface ReviewThreadLimits {
  threads: number;
  messagesPerThread: number;
  totalMessages: number;
  /** Characters per message. */
  textLength: number;
  authorLength: number;
  urlLength: number;
  labelLength: number;
}

/** A hyperlink as supplied by the host. */
export type HyperlinkInput =
  | { kind: "external"; url: string; label?: string }
  | { kind: "internal"; sheetId: string; address: string; label?: string };

/** A hyperlink after validation. An external URL is returned in its normalised form. */
export type HyperlinkRecord =
  | { kind: "external"; url: string; label?: string }
  | { kind: "internal"; sheetId: string; address: string; label?: string };

export interface FormulaWatch {
  id: string;
  anchor: { sheetId: string; sheetName: string; address: string };
  /** Reason the watch no longer points at a live cell, such as "missing-sheet". */
  orphaned?: string;
  [extra: string]: unknown;
}

/** Validates one data validation rule against a value. Returns true if the value is allowed; custom rules always return false here, so check them with checkDataValidationBatch. */
export function validateDataValidationRule(
  rule: DataValidationRule,
  value: unknown,
  datePolicy?: DatePolicy,
): boolean;

/** Returns normalised copies of the rules: ranges in upper case, unknown fields dropped. Throws if any rule is invalid or IDs repeat. */
export function normalizeDataValidationRules(
  validations: DataValidationRule[],
): DataValidationRule[];

/** Returns the data validation rules on a sheet whose range covers the cell. Returns an empty array when none do. Throws if the address is invalid. */
export function findDataValidationsInRange(
  workbook: Workbook,
  sheetId: string,
  cellAddress: string,
): DataValidationRule[];

/** Checks the target cells against their rules. Returns the number of issues approved (0 when nothing fails) and never changes the workbook; throws DataValidationError when a check fails and no approval covers it. */
export function checkDataValidationBatch(
  book: Workbook,
  targets: ValidationTarget[],
  validationApproval?: ValidationApproval,
): number;

/** Turns a DataValidationError that holds only warnings into an approval for checkDataValidationBatch. Throws if the error has no issues, or if any issue is a stop or could not be evaluated. */
export function assertNoStopValidationIssues(
  error: DataValidationErrorInstance,
): ValidationApproval;

/** Creates a DataValidationError for a list of issues. The message lists the first five issues. */
export const DataValidationError: new (
  workbook: Workbook,
  issues: DataValidationIssue[],
) => DataValidationErrorInstance;

/** Checks every value on a sheet against its rules, or on every sheet when no sheet ID is given. Returns a report and does not change the workbook. */
export function summarizeValidationInspection(
  workbook: Workbook,
  sheetId?: string,
): ValidationInspectionReport;

/** Serialises a report from summarizeValidationInspection as JSON text. Throws if the workbook or report has changed since the report was made. */
export function resolveValidationInspection(
  workbook: Workbook,
  report: ValidationInspectionReport,
): string;

/** Normalises conditional formatting rules and returns new copies. Throws if a rule is invalid, an ID or priority repeats, or a colour is not six-digit hex. */
export function normalizeConditionalRules(
  rules: ConditionalRule[],
): ConditionalRule[];

/** Returns a function that evaluates a sheet's conditional formats for one cell, and an errors() method that lists problems found so far. The calculator must be the callable from getWorkbookCalculation. */
export function evaluateConditionalFormatRules(
  sheet: Sheet,
  calculator: ConditionalCalculator,
  cells?: Record<string, Cell>,
): ConditionalFormatEvaluator;

/** Returns a preview of replacing a sheet's conditional formats with the given rules, and does not change the workbook. Throws if the sheet is missing, protected or the workbook is read-only. */
export function updateConditionalFormatRules(
  workbook: Workbook,
  sheetId: string,
  rules: ConditionalRule[],
  options?: RuleEditOptions,
): ConditionalFormatPreview;

/** Returns a new workbook with the previewed rules applied, or the same workbook if nothing changed. Throws if the preview is stale or the rules changed since it was made. */
export function assertConditionalFormatPreviewCurrent(
  workbook: Workbook,
  preview: ConditionalFormatPreview,
  options?: RuleEditOptions,
): Workbook;

/** Returns a copy of the rules sorted by priority; rules without a priority follow, last listed first. The input array is not changed. */
export function sortConditionalRulesByPriority(
  rules: ConditionalRule[],
): ConditionalRule[];

/** The conditional rule kinds the model supports, in their defined order. */
export const conditionalRuleKinds: readonly ConditionalRuleKind[];

/** Previews a defined-name create, update, rename or delete, including formulas it would rewrite. Does not change the workbook; throws if the operation is invalid or would break a dependent formula. */
export function applyDefinedNameOperation(
  workbook: Workbook,
  operation: DefinedNameOperation,
  options?: RuleEditOptions,
): DefinedNamePreview;

/** Returns the workbook with a defined-name preview applied, or throws if the workbook changed after the preview was made. */
export function assertNamePreviewCurrent(
  workbook: Workbook,
  preview: DefinedNamePreview,
  options?: RuleEditOptions,
): Workbook;

/** Replaces defined names in a formula with the text they refer to, and returns the expanded formula. Unknown names are left as they are. Throws "#CYCLE!" for a circular name and "#BLOCKED!" when canReadSheet refuses a sheet. */
export function resolveDefinedNameValue(
  formulaText: string,
  workbook: Workbook,
  currentSheetId: string,
  activeNameKeys?: string[],
  canReadSheet?: (sheetId: string) => boolean,
): string;

/** Adds a message to a review thread and returns a new workbook. Throws if the thread is missing, resolved, orphaned or read-only, or the text is empty or too long. */
export function addReviewThreadReply(
  workbook: Workbook,
  reviewThreadId: string,
  replyText: string,
  authorName?: string,
  options?: ReviewEditOptions,
): Workbook;

/** Resolves or reopens a review thread and returns a new workbook. Returns the same workbook when the thread is already in that state; throws if the thread is missing or the workbook is read-only. */
export function setReviewThreadResolved(
  workbook: Workbook,
  reviewThreadId: string,
  shouldResolve: boolean,
  authorName?: string,
  options?: ReviewEditOptions,
): Workbook;

/** Size limits for review threads, messages and hyperlinks. */
export const reviewThreadLimits: Readonly<ReviewThreadLimits>;

/** Checks a hyperlink and returns its normalised form. Accepts https, http and single-address mailto links, and worksheet links; throws for anything else. */
export function validateHyperlinkRecord(
  hyperlink: HyperlinkInput,
): HyperlinkRecord;

/** Sets or replaces the hyperlink on a cell and returns a new workbook. The revision and events are not changed. Throws if the sheet or address is invalid, the link is refused, or the workbook is read-only. */
export function setCellHyperlink(
  workbook: Workbook,
  sheetId: string,
  address: string,
  hyperlink: HyperlinkInput,
  options?: RuleEditOptions,
): Workbook;

/** Checks formula watches and returns normalised copies. A watch whose sheet is missing is kept and marked orphaned. Throws for duplicate IDs or two active watches on one cell. */
export function validateFormulaWatches(
  formulaWatches: FormulaWatch[] | undefined,
  workbook: Workbook,
): FormulaWatch[];
